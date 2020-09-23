Attribute VB_Name = "Utils"

Sub RestoreColor(frm As Form, clr As String)
Dim ctrl As Control
For Each ctrl In frm.Controls
If TypeOf ctrl Is Label Then
   ctrl.ForeColor = clr
End If
Next ctrl
End Sub

Sub UnloadAll()
'Unloads all the forms

Dim frm As Form
For Each frm In Forms
Unload frm
Next frm
End
End Sub

Sub FormMove(Form As Form)
'Will let you move a form if it
'doesn't have a titlebar
'best used in form or label mousedown event

Dim Ret&
ReleaseCapture
Ret& = SendMessage(Form.hWnd, &H112, &HF012, 0)
End Sub

Sub RollForm(frm As Form, upPos As Integer, downPos As Integer)
'Rolls a form up or down

If frm.Height = upPos Then
   Do
      frm.Height = frm.Height + 1
      DoEvents
   Loop Until frm.Height = downPos
Else
   Do
      frm.Height = frm.Height - 1
      DoEvents
   Loop Until frm.Height = upPos
End If
End Sub

Sub StayOnTop(frm As Form)
'duh

Dim KeepOnTop As Long
KeepOnTop = SetWindowPos(frm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
End Sub

Public Function FileExists(File As String) As Boolean
    On Error Resume Next
    If FileLen(File) > 0& Then
        If Err = 0 Then FileExists = True
    Else
       FileExists = False
    End If
End Function

Function RandomNumber(finished)
Randomize
RandomNumber = Int((Val(finished) * Rnd) + 1)
End Function

Public Function ReplaceString(MyString As String, ToFind As String, ReplaceWith As String) As String
    Dim Spot As Long, NewSpot As Long, LeftString As String
    Dim RightString As String, NewString As String
    Spot& = InStr(LCase(MyString$), LCase(ToFind))
    NewSpot& = Spot&
    Do
        If NewSpot& > 0& Then
            LeftString$ = Left(MyString$, NewSpot& - 1)
            If Spot& + Len(ToFind$) <= Len(MyString$) Then
                RightString$ = Right(MyString$, Len(MyString$) - NewSpot& - Len(ToFind$) + 1)
            Else
                RightString = ""
            End If
            NewString$ = LeftString$ & ReplaceWith$ & RightString$
            MyString$ = NewString$
        Else
            NewString$ = MyString$
        End If
        Spot& = NewSpot& + Len(ReplaceWith$)
        If Spot& > 0 Then
            NewSpot& = InStr(Spot&, LCase(MyString$), LCase(ToFind$))
        End If
    Loop Until NewSpot& < 1
    ReplaceString$ = NewString$
End Function

Sub StayNotOnTop(the As Form)
Dim SetWinOnTop  As Long
SetWinOnTop = SetWindowPos(the.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
End Sub

Public Function GetFromINI(Section As String, Key As String, Directory As String) As String
   Dim strBuffer As String
   strBuffer = String(750, Chr(0))
   Key$ = LCase$(Key$)
   GetFromINI$ = Left(strBuffer, GetPrivateProfileString(Section$, ByVal Key$, "", strBuffer, Len(strBuffer), Directory$))
End Function

Public Sub WriteToINI(Section As String, Key As String, KeyValue As String, Directory As String)
    Call WritePrivateProfileString(Section$, UCase$(Key$), KeyValue$, Directory$)
End Sub

Sub OpenEXE(FileName$)
Dim File As Double
File = Shell(FileName$, 1): NoFreeze% = DoEvents()
End Sub

Public Sub PlayMIDI(MIDIFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(MIDIFile$)
    If SafeFile$ <> "" Then
        Call mciSendString("play " & MIDIFile$, 0&, 0, 0)
    End If
End Sub

Public Sub StopMIDI(MIDIFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(MIDIFile$)
    If SafeFile$ <> "" Then
        Call mciSendString("stop " & MIDIFile$, 0&, 0, 0)
    End If
End Sub

Public Sub Playwav(WavFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(WavFile$)
    If SafeFile$ <> "" Then
        Call sndPlaySound(WavFile$, SND_FLAG)
    End If
End Sub

Sub Pause(Duration As Long)
'duh
    
Dim Current As Long
Current = Timer
Do Until Timer - Current >= Duration
    DoEvents
Loop
End Sub

Sub ErrMsg(msg As String)
MsgBox msg, 64, "Napster TooLz 1.0 By Skew & Jaze"
End Sub
