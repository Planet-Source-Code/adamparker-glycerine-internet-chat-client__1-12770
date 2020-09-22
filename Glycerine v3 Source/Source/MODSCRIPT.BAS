Attribute VB_Name = "modScript"
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'Some subs were taken from dos32
'Big thanks to Cim from Cimstudios.net
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1
Public Const GW_CHILD = 5
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1

Public Const LB_GETCOUNT = &H18B
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185

Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT

Public Const SW_HIDE = 0
Public Const SW_SHOW = 5

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_MENU = &H12
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
Public Const VK_SPACE = &H20
Public Const VK_UP = &H26

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
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Type POINTAPI
        X As Long
        Y As Long
End Type
Public Function GetFromINI(Section As String, Key As String, Directory As String) As String
   Dim strBuffer As String
   strBuffer = String(750, Chr(0))
   Key$ = LCase$(Key$)
   GetFromINI1$ = Left(strBuffer, GetPrivateProfileString(Section$, ByVal Key$, "", strBuffer, Len(strBuffer), Directory$))
   GetFromINI$ = DecryptA(GetFromINI1$)
End Function
Public Sub WriteToINI(Section As String, Key As String, KeyValue As String, Directory As String)
    KeyValue1$ = EncryptA(KeyValue$)
    Call WritePrivateProfileString(Section$, UCase$(Key$), KeyValue1$, Directory$)
End Sub

Public Sub Playwav(WavFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(WavFile$)
    If SafeFile$ <> "" Then
        Call sndPlaySound(WavFile$, SND_FLAG)
    End If
End Sub
Public Function EncryptA(txt As String)
    itz = ChrW(Random(25, 50)) + ReplaceString(txt, " ", "·   ") + ChrW(Random(25, 50))
    For i = 1 To Len(itz)
    Mi = Mid$(itz, i, 1)
    aa = Asc(Mi)
    ab = (aa + i) - 8
    ap = Chr$(ab)
    ch = ch & ap
    Next i
    EncryptA = ch
End Function
Public Function DecryptA(txt)
    itz = txt
    For i = 1 To Len(itz)
    Mi = Mid$(itz, i, 1)
    aa = Asc(Mi)
    ab = (aa - i) + 8
    ap = Chr$(ab)
    ch$ = ch$ & ap
    Next i
    If ch$ = Empty Then
        DecryptA = ch$
    Else
        ch$ = Right(ch$, Len(ch$) - 1)
        ch$ = Left(ch$, Len(ch$) - 1)
        ch$ = ReplaceString(ch$, "·   ", " ")
        DecryptA = ch$
    End If
End Function
Public Function Random(intFrom As Integer, intTo As Integer)
    Randomize
    Result = Int((intTo * Rnd) + intFrom)
    Random = Result
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
Public Function ReverseString(MyString As String) As String
    Dim TempString As String, StringLength As Long
    Dim Count As Long, NextChr As String, NewString As String
    TempString$ = MyString$
    StringLength& = Len(TempString$)
    Do While Count& <= StringLength&
        Count& = Count& + 1
        NextCr$ = Mid$(TempString$, Count&, 1)
        NewString$ = NextChr$ & NewString$
    Loop
    ReverseString$ = NewString$
End Function

Sub SaveText(txtSave, PathA As String)
    Dim TextString As String
    On Error Resume Next
    TextString$ = txtSave
    Open PathA$ For Output As #1
    Print #1, TextString$
    Close #1
End Sub
Public Sub Loadlistbox(Directory As String, TheList)
    Dim MyString As String
    On Error Resume Next
Dim a As Variant
Dim b As Variant
a = 1
Open Directory$ For Input As a
While (EOF(a) = False)
Line Input #a, b
TheList.AddItem b
Wend
Close a
End Sub

Public Sub SaveListBox(Directory As String, TheList)
    Dim SaveList As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveList& = 0 To TheList.ListCount - 1
        Print #1, TheList.List(SaveList&)
    Next SaveList&
    Close #1
End Sub
Public Sub FormOnTop(FormName As Form)
    Call SetWindowPos(FormName.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub

Public Sub FormNotOnTop(FormName As Form)
    Call SetWindowPos(FormName.hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub
Function EliteTalker(word$)
Made$ = ""
For Q = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, Q, 1)
    Leet$ = ""
    X = Int(Rnd * 3 + 1)
    If letter$ = "a" Then
    If X = 1 Then Leet$ = "â"
    If X = 2 Then Leet$ = "å"
    If X = 3 Then Leet$ = "ä"
    End If
    If letter$ = "b" Then Leet$ = "b"
    If letter$ = "c" Then Leet$ = "ç"
    If letter$ = "d" Then Leet$ = "d"
    If letter$ = "e" Then
    If X = 1 Then Leet$ = "ë"
    If X = 2 Then Leet$ = "ê"
    If X = 3 Then Leet$ = "é"
    End If
    If letter$ = "i" Then
    If X = 1 Then Leet$ = "ì"
    If X = 2 Then Leet$ = "ï"
    If X = 3 Then Leet$ = "î"
    End If
    If letter$ = "j" Then Leet$ = ",j"
    If letter$ = "n" Then Leet$ = "ñ"
    If letter$ = "o" Then
    If X = 1 Then Leet$ = "ô"
    If X = 2 Then Leet$ = "ð"
    If X = 3 Then Leet$ = "õ"
    End If
    If letter$ = "s" Then Leet$ = "š"
    If letter$ = "t" Then Leet$ = "†"
    If letter$ = "u" Then
    If X = 1 Then Leet$ = "ù"
    If X = 2 Then Leet$ = "û"
    If X = 3 Then Leet$ = "ü"
    End If
    If letter$ = "w" Then Leet$ = "vv"
    If letter$ = "y" Then Leet$ = "ÿ"
    If letter$ = "0" Then Leet$ = "Ø"
    If letter$ = "A" Then
    If X = 1 Then Leet$ = "Å"
    If X = 2 Then Leet$ = "Ä"
    If X = 3 Then Leet$ = "Ã"
    End If
    If letter$ = "B" Then Leet$ = "ß"
    If letter$ = "C" Then Leet$ = "Ç"
    If letter$ = "D" Then Leet$ = "Ð"
    If letter$ = "E" Then Leet$ = "Ë"
    If letter$ = "I" Then
    If X = 1 Then Leet$ = "Ï"
    If X = 2 Then Leet$ = "Î"
    If X = 3 Then Leet$ = "Í"
    End If
    If letter$ = "N" Then Leet$ = "Ñ"
    If letter$ = "O" Then Leet$ = "Õ"
    If letter$ = "S" Then Leet$ = "Š"
    If letter$ = "U" Then Leet$ = "Û"
    If letter$ = "W" Then Leet$ = "VV"
    If letter$ = "Y" Then Leet$ = "Ý"
    If letter$ = "`" Then Leet$ = "´"
    If letter$ = "!" Then Leet$ = "¡"
    If letter$ = "?" Then Leet$ = "¿"
    If Len(Leet$) = 0 Then Leet$ = letter$
    Made$ = Made$ & Leet$
Next Q
EliteTalker = Made$
End Function
Public Function LineCount(MyString As String) As Long
    Dim Spot As Long, Count As Long
    If Len(MyString$) < 1 Then
        LineCount& = 0&
        Exit Function
    End If
    Spot& = InStr(MyString$, Chr(13))
    If Spot& <> 0& Then
        LineCount& = 1
        Do
            Spot& = InStr(Spot + 1, MyString$, Chr(13))
            If Spot& <> 0& Then
                LineCount& = LineCount& + 1
            End If
        Loop Until Spot& = 0&
    End If
    LineCount& = LineCount& + 1
End Function
Public Function LineFromString(MyString As String, Line As Long) As String
    Dim theline As String, Count As Long
    Dim FSpot As Long, LSpot As Long, DoIt As Long
    Count& = LineCount(MyString$)
    If Line& > Count& Then
        Exit Function
    End If
    If Line& = 1 And Count& = 1 Then
        LineFromString$ = MyString$
        Exit Function
    End If
    If Line& = 1 Then
        theline$ = Left(MyString$, InStr(MyString$, Chr(13)) - 1)
        theline$ = ReplaceString(theline$, Chr(13), "")
        theline$ = ReplaceString(theline$, Chr(10), "")
        LineFromString$ = theline$
        Exit Function
    Else
        FSpot& = InStr(MyString$, Chr(13))
        For DoIt& = 1 To Line& - 1
            LSpot& = FSpot&
            FSpot& = InStr(FSpot& + 1, MyString$, Chr(13))
        Next DoIt
        If FSpot = 0 Then
            FSpot = Len(MyString$)
        End If
        theline$ = Mid(MyString$, LSpot&, FSpot& - LSpot& + 1)
        theline$ = ReplaceString(theline$, Chr(13), "")
        theline$ = ReplaceString(theline$, Chr(10), "")
        LineFromString$ = theline$
    End If
End Function
Function TrimSpaces(Text)
    If InStr(Text, " ") = 0 Then
        TrimSpaces = Text
        Exit Function
    End If
    For TrimSpace = 1 To Len(Text)
    thechar$ = Mid(Text, TrimSpace, 1)
    thechars$ = thechars$ & thechar$
    If thechar$ = " " Then
        thechars$ = Mid(thechars$, 1, Len(thechars$) - 1)
    End If
    Next TrimSpace
    TrimSpaces = thechars$
End Function
Function TrimFirstSpaces(strString) As String
    strEndString = strString
    For intBuffer = 1 To Len(strString)
        If Mid(strString, intBuffer, 1) = " " Then
            strEndString = Right(strEndString, Len(strEndString) - 1)
        Else
            TrimFirstSpaces = strEndString
            Exit Function
        End If
    Next intBuffer
End Function
Sub LoadText(txtLoad, PathA As String)
    Dim TextString As String
    On Error Resume Next
    Open PathA$ For Input As #1
    TextString$ = Input(LOF(1), #1)
    Close #1
    txtLoad.Text = TextString$
End Sub
Function DoScript(strBuffer As String)
On Error GoTo err:
    strTheScript = strBuffer
    Dim intBuffer As Long, strLine As String, strType As String, strInput As String, strString As String, strIF As String, variable1 As String, variable2 As String
    For intBuffer = 1 To LineCount(strBuffer)
        strString = LineFromString(strBuffer, intBuffer)
        If Not Left(strString, 1) = ";" Then
            If Not Len(strString) = 0 Then
                If Not Right(strString, 1) = ";" Then
                    MsgBox "Script Error:" & vbNewLine & "Line " & intBuffer, vbExclamation + vbOKOnly, "Glycerine: Script Error"
                    Exit Function
                End If
                strString = Left(strString, Len(strString) - 1)
                strString = TrimFirstSpaces(strString)
                If Left(strString, 6) = "server" Then
                    intMarker = InStr(strString, "(")
                    If intMarker = 0 Then
                        MsgBox "Script Error:" & vbNewLine & "Line " & intBuffer, vbExclamation + vbOKOnly, "Glycerine: Script Error"
                        Exit Function
                    Else
                    intMarkerA = InStrRev(strString, ")")
                    server = Mid(strString, intMarker + 1, intMarkerA - intMarker - 1)
                    frmConnect.cmbServer.AddItem (server)
                    End If
                    End If
                If Left(strString, 7) = "dserver" Then
                    intMarker = InStr(strString, "(")
                    If intMarker = 0 Then
                        MsgBox "Script Error:" & vbNewLine & "Line " & intBuffer, vbExclamation + vbOKOnly, "Glycerine: Script Error"
                        Exit Function
                    Else
                    intMarkerA = InStrRev(strString, ")")
                    server = Mid(strString, intMarker + 1, intMarkerA - intMarker - 1)
                    frmConnect.cmbServer.Text = server
                    End If
                    End If
                    If Left(strString, 6) = "append" Then
                    Dim strAppend As String
                    intMarker = InStr(strString, "(")
                    If intMarker = 0 Then
                        MsgBox "Script Error:" & vbNewLine & "Line " & intBuffer, vbExclamation + vbOKOnly, "Glycerine: Script Error"
                        Exit Function
                    Else
                        intMarkerA = InStrRev(strString, ")")
                        If intMarkerA = 0 Then
                            MsgBox "Script Error:" & vbNewLine & "Line " & intBuffer, vbExclamation + vbOKOnly, "Glycerine: Script Error"
                            Exit Function
                        End If
                        strString = Mid(strString, intMarker + 1, intMarkerA - intMarker - 1)
                        intMarkerB = InStr(strString, ",")
                        If intMarkerB = 0 Then
                            MsgBox "Script Error:" & vbNewLine & "Line " & intBuffer, vbExclamation + vbOKOnly, "Glycerine: Script Error"
                            Exit Function
                        End If
                        strAppend = Right(strString, Len(strString) - intMarkerB)
                        strFile = Left(strString, intMarkerB - 1)
                        If Not Len(Dir(App.Path & "\" & strFile)) > 0 Then
                            MsgBox "Script Error:" & vbNewLine & "Line " & intBuffer, vbExclamation + vbOKOnly, "Glycerine: Script Error"
                            Exit Function
                        Else
                            Open App.Path & "\" & strFile For Append As #1
                            Print #1, ReplaceString(strAppend, "$nl", vbNewLine)
                            Close #1
                        End If
                    End If
                End If
            End If
        End If
    Next intBuffer
Exit Function
err:
MsgBox "Error: " & err.Description, vbOKOnly, err.Number
End Function
Public Sub TimeOut(Duration As Long)
    Dim Current As Long
    Current = Timer
    Do Until Timer - Current >= Duration
    DoEvents
    Loop
End Sub
Public Function FileExists(sFileName As String) As Boolean
    If Len(sFileName$) = 0 Then
        FileExists = False
        Exit Function
    End If
    If Len(Dir$(sFileName$)) Then
        FileExists = True
    Else
        FileExists = False
    End If
End Function


