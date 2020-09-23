Attribute VB_Name = "napBasBySkew"
'**************************************
'*                                    *
'* Napster Bas by Skew & Jaze         *
'* http://go.to/GSoftware/            *
'* Ghetto Software Inc.               *
'*                                    *
'* If this bas is to be used in       *
'* another program the authors        *
'* must be given credit               *
'* somewhere in the program.          *
'*                                    *
'* Skew - xskewx@hotmail.com          *
'* Jaze - jaze_philly@hotmail.com     *
'*                                    *
'* Notes:                             *
'* - This bas was a bitch to make,    *
'* mainly because of the way napster  *
'* categorizes it's windows and uses  *
'* that damn SysTab control.          *
'* This is the reason most of the     *
'* functions will do whatever they    *
'* have to do with the first chatroom *
'* after the "PRIVATE" chatroom.      *
'* It is possible to acccess the      *
'* other chatrooms but it is just     *
'* annoying and there's no need to.   *
'* -If you are going to use one of    *
'* the functions that deal with       *
'* chatrooms remeber that it only     *
'* deals with the first available     *
'* chatroom after the "PRIVATE" chat. *
'* -Some ideas were taken from        *
'* Dos - www.dosfx.com                *
'*                                    *
'**************************************

'API Declarations
'A lot of these aren't used but are in here anyway because I use them at other times

Option Explicit

Public Declare Function GetTopWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowThreadProcessID Lib "user32" Alias "GetWindowThreadProcessId" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hwndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Public Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source&, ByVal dest&, ByVal nCount&)
Public Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Public Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source As String, ByVal dest As Long, ByVal nCount&)
Public Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Public Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hWnd&)
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function sendmessagebynum& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function iswindowenabled Lib "user32" Alias "IsWindowEnabled" (ByVal hWnd As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Public Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Public Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long)
Public Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal cmd As Long) As Long
Public Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Global Contants
Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1

Public Const LB_GETITEMDATA = &H199
Public Const LB_GETCOUNT = &H18B
Public Const LB_ADDSTRING = &H180
Public Const LB_DELETESTRING = &H182
Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_GETCURSEL = &H188
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SELECTSTRING = &H18C
Public Const LB_SETCOUNT = &H1A7
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185
Public Const LB_INSERTSTRING = &H181

Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT

Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_MAX = 5
Public Const GW_OWNER = 4
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_HIDE = 0
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1

Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE

Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_MENU = &H12
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
Public Const VK_SPACE = &H20
Public Const VK_UP = &H26

Public Const MF_APPEND = &H100&
Public Const MF_DELETE = &H200&
Public Const MF_CHANGE = &H80&
Public Const MF_ENABLED = &H0&
Public Const MF_DISABLED = &H2&
Public Const MF_REMOVE = &H1000&
Public Const MF_POPUP = &H10&
Public Const MF_STRING = &H0&
Public Const MF_UNCHECKED = &H0&
Public Const MF_CHECKED = &H8&
Public Const MF_GRAYED = &H1&
Public Const MF_BYPOSITION = &H400&
Public Const MF_BYCOMMAND = &H0&

Public Const GWW_HINSTANCE = (-6)
Public Const GWW_ID = (-12)
Public Const GWL_STYLE = (-16)

Public Const WM_CHAR = &H102
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MOVE = &HF012
Public Const WM_SETTEXT = &HC
Public Const WM_SYSCOMMAND = &H112

Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000

Public Const ENTER_KEY = 13

Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Public Type POINTAPI
   X As Long
   Y As Long
End Type

Public Function FindNapster() As Long
'Can easily be changed to detect any version of napster
'Thanks to Ziegs for pointing out a fix for this function

Dim nap As Long
nap = FindWindow("NAPSTER", "Napster v2.0 BETA 7")
If nap <> 0& Then
   FindNapster = nap
Else
   nap = FindWindow("NAPSTER", "Napster v2.0 BETA 8")
   If nap <> 0& Then
      FindNapster = nap
   Else
      nap = FindWindow("NAPSTER", "Napster v2.0 BETA 8.24")
      If nap <> 0& Then
         FindNapster = nap
      Else
         FindNapster = FindWindow("NAPSTER", "Napster v2.0 BETA 9")
      End If
   End If
End If
End Function

Public Function FindPrivate() As Long
Dim nap As Long
Dim chat As Long
On Error Resume Next
nap = FindNapster
chat = FindWindowEx(nap, 0&, "#32770", vbNullString)
FindPrivate = FindWindowEx(chat, 0&, "#32770", vbNullString)
End Function

Public Sub PrivateCommand(Text As String)
'This should be used to send commands
'to the main "PRIVATE" chat

Dim nap, win, txt As Long
On Error Resume Next
nap = FindNapster
win = FindWindowEx(nap, 0&, "#32770", vbNullString)
txt = FindWindowEx(win, 0&, "#32770", "PRIVATE")
txt = FindWindowEx(txt, 0&, "RICHEDIT", vbNullString)
Call SendMessageByString(txt, WM_SETTEXT, 0&, Text)
Call SendMessageLong(txt, WM_CHAR, ENTER_KEY, 0&)

End Sub

Public Sub ChatSend(Text As String)
'Sends chat to the first chatroom
'after the "PRIVATE" chat
Dim nap, win, chat, txtbox As Long
nap = FindNapster
win = FindWindowEx(nap, 0&, "#32770", vbNullString)
chat = FindWindowEx(win, 0&, "#32770", "PRIVATE")
chat = FindWindowEx(win, chat, "#32770", vbNullString)
txtbox = FindWindowEx(chat, 0&, "RICHEDIT", vbNullString)
Call SendMessageByString(txtbox, WM_SETTEXT, 0&, Text)
Call SendMessageLong(txtbox, WM_CHAR, ENTER_KEY, 0&)
End Sub

Public Sub RunMenuByString(menu As String)
'Will rum the menu item by
'the string you send it
'Ex. The about napster menu would be
'RunMenuByString("&About Napster")
'Must take into account the underlined letter

Dim nap, napmenu, MenuCount, lookfor As Long
Dim smenu, scount, looksub, sid As Long
Dim sstring As String
On Error Resume Next
nap = FindNapster
napmenu = GetMenu(nap)
MenuCount = GetMenuItemCount(napmenu)
For lookfor = 0& To MenuCount - 1
   smenu = GetSubMenu(napmenu, lookfor)
   scount = GetMenuItemCount(smenu)
   For looksub = 0 To scount - 1
      sid = GetMenuItemID(smenu, looksub)
      sstring = String(100, " ")
      Call GetMenuString(smenu, sid, sstring, 100&, 1&)
      If InStr(LCase(sstring), LCase(menu)) Then
         Call SendMessageLong(nap, WM_COMMAND, sid, 0&)
         Exit Sub
      End If
   Next looksub
Next lookfor
End Sub

Public Sub Search(artist As String, title As String, maxresults As String)
'This will do a search for you
'in the search thing
Dim nap, win, txtartist, txttitle, txtmaxresults, findit As Long
nap = FindNapster
win = FindWindowEx(nap, 0&, "#32770", vbNullString)
win = FindWindowEx(nap, win, "#32770", vbNullString)
win = FindWindowEx(nap, win, "#32770", vbNullString)
txtartist = FindWindowEx(win, 0&, "EDIT", vbNullString)
Call SendMessageByString(txtartist, WM_SETTEXT, 0&, artist)
txttitle = FindWindowEx(win, txtartist, "EDIT", vbNullString)
Call SendMessageByString(txttitle, WM_SETTEXT, 0&, title)
txtmaxresults = FindWindowEx(win, txttitle, "EDIT", vbNullString)
Call SendMessageByString(txtmaxresults, WM_SETTEXT, 0&, maxresults)
findit = FindWindowEx(win, 0&, "Button", "Find it!")
Call ClickButton(findit)
End Sub

Public Sub InstantMessage(person As String, Message As String)
'Will instant message a user with the message

Call PrivateCommand("/msg " & person & " " & Message)
End Sub

Public Sub ClickButton(mButton As Long)
'Clicks the button with the handle passed in
        
Call SendMessage(mButton, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessage(mButton, WM_KEYUP, VK_SPACE, 0&)
End Sub

Public Function IsUserOnline() As Boolean
'Will tell you if the user is
'connected to the napster server
Dim nap, systab As Long
Dim txt As String
nap = FindNapster
systab = FindWindowEx(nap, 0&, "msctls_statusbar32", vbNullString)
txt = Left(GetText(systab), 6)
If txt = "Online" Then
   IsUserOnline = True
Else
   IsUserOnline = False
End If
End Function

Private Function GetText(ByVal windowhandle As Long) As String
'Gets text of a window
Dim textlength As Long
Dim buffer As String
DoEvents
textlength = SendMessage(windowhandle, WM_GETTEXTLENGTH, 0&, 0&)
buffer = String(textlength, 0&)
Call SendMessageByString(windowhandle, WM_GETTEXT, textlength + 1, buffer)
GetText = buffer
End Function

Public Function UserSN() As String
'Gets the users SceenName
Dim nap, systab As Long
Dim txt As String
On Error Resume Next
nap = FindNapster
systab = FindWindowEx(nap, 0&, "msctls_statusbar32", vbNullString)
txt = GetText(systab)
UserSN = Mid(txt, (InStr(txt, "(")) + 1, (InStr(txt, ")") - 1) - (InStr(txt, "(")))
End Function

Public Function FindChatRoom() As Long
'Finds the handle of the first chatroom
'After the "Private" chat

Dim nap, win, chat As Long
DoEvents
nap = FindNapster
win = FindWindowEx(nap, 0&, "#32770", vbNullString)
chat = FindWindowEx(win, 0&, "#32770", "PRIVATE")
chat = FindWindowEx(win, chat, "#32770", vbNullString)
FindChatRoom = chat
End Function

Public Function GetChatText() As String
'Will get the last line of chat
'from the first chatroom after
'"PRIVATE" chat

Dim room, txt, chat As Long
Dim chattext As String
DoEvents
room = FindChatRoom
txt = FindWindowEx(room, 0&, "RICHEDIT", vbNullString)
chat = FindWindowEx(room, txt, "RICHEDIT", vbNullString)
chattext = GetText(chat)
GetChatText = chattext
End Function

Public Function LastChatLineWithSN() As String
'Formats the GetChatText into
'the last msg and screename

Dim chattext, thechar, thechars, TheChatText, lastline As String
Dim findchat, lastlen, findchar As Long

DoEvents
chattext = GetChatText
For findchar = 1 To Len(chattext)
   DoEvents
   thechar = Mid(chattext, findchar, 1)
   thechars = thechars & thechar
   If thechar = Chr(13) Then
      DoEvents
      TheChatText = Mid(thechars, 1, Len(thechars) - 1)
      thechars = ""
   End If
   DoEvents
Next findchar
lastlen = Val(findchar) - Len(thechars)
lastline = Mid(chattext, lastlen, Len(thechars))
DoEvents
LastChatLineWithSN = lastline
End Function

Public Function ValidChatLine() As String
'Gets a valid chatline from chat

Dim chat As String
chat = LastChatLineWithSN
If InStr(chat, "*") = 3 Then
   Do
      DoEvents
      chat = LastChatLineWithSN
   Loop Until InStr(chat, "*") <> 3
End If
ValidChatLine = chat
End Function

Public Function SNFromLastChatLine() As String
'Gets the screenname from
'the lastchatline without the chat

Dim chat As String
chat = ValidChatLine
SNFromLastChatLine = Mid(chat, 4, (InStr(chat, ">") - 4))
End Function

Public Function LastChatLine() As String
'Gets the text from the last
'chat line without screenname

Dim chat As String
chat = ValidChatLine
LastChatLine = Mid(chat, InStr(chat, ">") + 1, Len(chat) - InStr(chat, ">"))
End Function

Public Sub HideNapster(SW_Command As Integer)
'Will  hide or show the napster window
'depending the command
'hide = SW_HIDE  show = SW_SHOW

Dim nap As Long
nap = FindNapster
Call ShowWindow(nap, SW_Command)
End Sub

Private Sub SetText(Window As Long, Text As String)
'Sets the text in a window
'Used in clearchat

Call SendMessageByString(Window, WM_SETTEXT, 0&, Text)
End Sub

Public Function ChatRoomName() As String
'Gets the name of the first
'chatroom after "PRIVATE" Chat

Dim room, chat As Long
Dim name As String
On Error Resume Next
DoEvents
room = FindChatRoom
chat = FindWindowEx(room, 0&, "RICHEDIT", vbNullString)
chat = FindWindowEx(room, chat, "RICHEDIT", vbNullString)
chat = FindWindowEx(room, chat, "RICHEDIT", vbNullString)
name = GetText(chat)
name = Mid(name, 2, InStr(name, ":") - 2)
ChatRoomName = name
End Function

Public Function NumPeopleInChat() As String
'Get the number of people in a chatroom

Dim room, chat As Long
Dim numpeeps As String
room = FindChatRoom
chat = FindWindowEx(room, 0&, "RICHEDIT", vbNullString)
chat = FindWindowEx(room, chat, "RICHEDIT", vbNullString)
chat = FindWindowEx(room, chat, "RICHEDIT", vbNullString)
chat = FindWindowEx(room, chat, "RICHEDIT", vbNullString)
numpeeps = GetText(chat)
NumPeopleInChat = Mid(numpeeps, 2, InStr(numpeeps, "u") - 3)
End Function

Public Function ChatRoomInfo() As String
'Gets the info of the first
'chatroom after "PRIVATE" Chat

Dim room, chat As Long
Dim name As String
room = FindChatRoom
chat = FindWindowEx(room, 0&, "RICHEDIT", vbNullString)
chat = FindWindowEx(room, chat, "RICHEDIT", vbNullString)
chat = FindWindowEx(room, chat, "RICHEDIT", vbNullString)
name = GetText(chat)
ChatRoomInfo = name
End Function

Public Sub ClearPrivate()
'Easier way of clearing the
'"PRIVATE" chat

PrivateCommand ("/Clear")
End Sub

Public Sub ClearChat()
'Clears the first chatroom after
'"PRIVATE" chat

Dim nap, win, chat, edit As Long
nap = FindNapster
win = FindWindowEx(nap, 0&, "#32770", vbNullString)
chat = FindWindowEx(win, 0&, "#32770", "PRIVATE")
chat = FindWindowEx(win, chat, "#32770", vbNullString)
edit = FindWindowEx(chat, 0&, "RICHEDIT", vbNullString)
edit = FindWindowEx(chat, edit, "RICHEDIT", vbNullString)
Call SetText(edit, " ")
End Sub

Public Function ChatLineSN(TheChatLine As String) As String
'Gets the SN from a chatline

On Error Resume Next
ChatLineSN = Mid(TheChatLine, 4, (InStr(TheChatLine, ">") - 4))
End Function

Public Function ChatLineMsg(TheChatLine As String) As String
'Get the text from a chatline

On Error Resume Next
ChatLineMsg = Mid(TheChatLine, InStr(TheChatLine, ">") + 2, Len(TheChatLine) - InStr(TheChatLine, ">"))
End Function

Public Function FindIM() As Long
'Find the first active IM Window

Dim imwin As Long
Dim cap As String
imwin = FindWindow("#32770", vbNullString)
cap = GetCaption(imwin)
If Right(cap, 15) <> "Instant Message" Then
   Do
      DoEvents
      imwin = FindWindow("#32770", vbNullString)
      cap = GetCaption(imwin)
      If Left(cap, 15) = "Instant Message" Then
         FindIM = imwin
         Exit Function
      End If
   Loop Until imwin <> 0&
Else
   FindIM = imwin
   Exit Function
End If
FindIM = 0&
End Function

Public Function SNfromIM(imtext As String) As String
'Gets the sn from the im line in "IMText"

SNfromIM = Mid(imtext, 4, InStr(imtext, ">") - 4)
End Function

Public Function MsgFromIM(imtext As String) As String
'Gets the msg from the IM line in "IMText"

MsgFromIM = Mid(imtext, InStr(imtext, ">") + 2, Len(imtext) - (InStr(imtext, ">") + 1))
End Function

Public Function LastIMLineWithSN(ByVal IMHandle As Long) As String
'Gets the last IM line and returns it

Dim txbx, findchar, thechars, lastlen As Long
Dim imtext, lastline, thechar, TheChatText As String
txbx = FindWindowEx(IMHandle, 0&, "RICHEDIT", vbNullString)
txbx = FindWindowEx(IMHandle, txbx, "RICHEDIT", vbNullString)
imtext = GetText(txbx)
DoEvents
For findchar = 1 To Len(imtext)
   DoEvents
   thechar = Mid(imtext, findchar, 1)
   thechars = thechars & thechar
   If thechar = Chr(13) Then
      DoEvents
      TheChatText = Mid(thechars, 1, Len(thechars) - 1)
      thechars = ""
   End If
   DoEvents
Next findchar
lastlen = Val(findchar) - Len(thechars)
lastline = Mid(imtext, lastlen, Len(thechars))
DoEvents
LastIMLineWithSN = lastline
End Function

Private Function GetCaption(windowhandle As Long) As String
'Gets the caption of the window
    
Dim buffer As String
Dim textlength As Long
textlength = GetWindowTextLength(windowhandle)
buffer = String(textlength, 0&)
Call GetWindowText(windowhandle, buffer, textlength + 1)
GetCaption = buffer
End Function

Public Sub RespondIM(IMHandle As Long, msg As String)
'Respond an IM with msg

Dim txtbx As Long
DoEvents
txtbx = FindWindowEx(IMHandle, 0&, "RICHEDIT", vbNullString)
Call SendMessageByString(txtbx, WM_SETTEXT, 0&, msg)
Call SendMessageLong(txtbx, WM_CHAR, ENTER_KEY, 0&)
End Sub

Public Function GetPrivateText() As String
'Gets the text from the "PRIVATE" text

Dim chat, txt As Long
DoEvents
chat = FindPrivate
txt = FindWindowEx(chat, 0&, "RICHEDIT", vbNullString)
txt = FindWindowEx(chat, txt, "RICHEDIT", vbNullString)
GetPrivateText = GetText(txt)
End Function

Public Function LastPrivateLine()
'Gets the last line from the "PRIVATE" chat
Dim privtext, thechar, lastline, TheChatText As String
Dim findchar, thechars, lastlen As Long
DoEvents
privtext = GetPrivateText
For findchar = 1 To Len(privtext)
   DoEvents
   thechar = Mid(privtext, findchar, 1)
   thechars = thechars & thechar
   If thechar = Chr(13) Then
      DoEvents
      TheChatText = Mid(thechars, 1, Len(thechars) - 1)
      thechars = ""
   End If
   DoEvents
Next findchar
lastlen = Val(findchar) - Len(thechars)
lastline = Mid(privtext, lastlen, Len(thechars))
LastPrivateLine = lastline
End Function

Public Sub Command(cmd As String, UserorMsg As String, reason As String)
'Gets the reply after a command is sent

Call PrivateCommand("/" & cmd & " " & UserorMsg & " " & reason)
End Sub

Public Sub WaitForMsg()
'Gets the last msg from "PRIVATE" chat

Dim chat As String
Dim chatlen, chatlen2 As Integer
chatlen = Len(GetPrivateText)
Do
   DoEvents
   chatlen2 = Len(GetPrivateText)
Loop Until chatlen2 > chatlen
ErrMsg (LastPrivateLine)
End Sub

Function BallAnswer() As String
Dim num As Integer
num = RandomNumber(8)
If num = 1 Then BallAnswer = "Yes"
If num = 2 Then BallAnswer = "No"
If num = 3 Then BallAnswer = "Outlook Dim"
If num = 4 Then BallAnswer = "Unlikely"
If num = 5 Then BallAnswer = "Doubtful"
If num = 6 Then BallAnswer = "Probable"
If num = 7 Then BallAnswer = "Definitely"
If num = 8 Then BallAnswer = "Outlook Good"
End Function

Function FortuneAnswer() As String
Dim num As Integer
num = RandomNumber(10)
If num = 1 Then FortuneAnswer = "Today will be a good day"
If num = 2 Then FortuneAnswer = "Today will be a bad day"
If num = 3 Then FortuneAnswer = "Be careful of who you trust"
If num = 4 Then FortuneAnswer = "You will find a new love"
If num = 5 Then FortuneAnswer = "The stars are aligned in your favor"
If num = 6 Then FortuneAnswer = "Look where you least expect"
If num = 7 Then FortuneAnswer = "Trust the words of a friend"
If num = 8 Then FortuneAnswer = "Love is not on your side today"
If num = 9 Then FortuneAnswer = "Leave the one you are with"
If num = 10 Then FortuneAnswer = "your friends will always be there"
End Function
