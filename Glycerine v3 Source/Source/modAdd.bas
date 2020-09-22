Attribute VB_Name = "modAdd"

Sub AddChat(strData As String)
On Error Resume Next
  strUser = Left$(strData, InStr(strData, ":") - 1)
  strText = Right(strData, Len(strData) - InStr(strData, ":"))
  
  strTime = strText
  strTime = TrimSpaces(strTime)
  If Mid(strTime, 1, 6) Then
  strTime = Mid(strTime, 7)
  strTime = Left$(strTime, InStr(strTime, "~") - 1)
  strColor = strTime
  End If
  If strColor = "" Then strColor = "0"
  strText = Mid(strText, 7)
  strText = Right(strData, Len(strText) - InStr(strText, "~"))
  strText = "    " & strText
  
  strUser = "<" & strUser & ">"
  If mdiMain.mnuFileTime.Checked = True Then
  strUser = " [" & Time & "]  " & strUser
  End If
Dim lngLastLen As Long
  frmMain.txtText.SelStart = Len(frmMain.txtText.Text)
  frmMain.txtText.SelLength = 0
  frmMain.txtText.SelText = strUser
  frmMain.txtText.SelStart = Len(frmMain.txtText.Text) - (Len(strUser))
  frmMain.txtText.SelLength = Len(strUser) + 4
  frmMain.txtText.SelColor = vbBlue
  frmMain.txtText.SelFontSize = 8
  frmMain.txtText.SelBold = True
  frmMain.txtText.SelUnderline = False
  frmMain.txtText.SelItalic = False
  lngLastLen& = Len(frmMain.txtText.Text)
  frmMain.txtText.SelStart = lngLastLen&
  frmMain.txtText.SelLength = 0
  frmMain.txtText.SelRTF = "   " & strText & vbNewLine
  frmMain.txtText.SelStart = lngLastLen&
  frmMain.txtText.SelLength = Len(frmMain.txtText.Text) - lngLastLen&
  frmMain.txtText.SelColor = strColor
  frmMain.txtText.SelFontSize = 8
  frmMain.txtText.SelBold = True
  frmMain.txtText.SelUnderline = False
  frmMain.txtText.SelItalic = False
  frmMain.txtText.SelHangingIndent = 1400
  frmMain.txtText.SelStart = Len(frmMain.txtText.Text)
  frmMain.txtText.SelLength = 0
End Sub

Sub AddChat2(strData As String)
On Error Resume Next
  strUser = Left$(strData, InStr(strData, ":") - 1)
  strText = Right(strData, Len(strData) - InStr(strData, ":"))
  
  strTime = strText
  strTime = TrimSpaces(strTime)
  If Mid(strTime, 1, 6) Then
  strTime = Mid(strTime, 7)
  strTime = Left$(strTime, InStr(strTime, "~") - 1)
  strColor = strTime
  End If
  If strColor = "" Then strColor = "0"
  strText = Mid(strText, 7)
  strText = Right(strData, Len(strText) - InStr(strText, "~"))
  strText = "    " & strText
  
  strUser = "<" & strUser & ">"
  If mdiMain.mnuFileTime.Checked = True Then
  strUser = " [" & Time & "]  " & strUser
  End If
  Dim lngLastLen As Long
  frmMain.txtText.SelStart = Len(frmMain.txtText.Text)
  frmMain.txtText.SelLength = 0
  frmMain.txtText.SelText = strUser
  frmMain.txtText.SelStart = Len(frmMain.txtText.Text) - (Len(strUser))
  frmMain.txtText.SelLength = Len(strUser) + 4
  frmMain.txtText.SelColor = vbRed
  frmMain.txtText.SelFontSize = 8
  frmMain.txtText.SelBold = True
  frmMain.txtText.SelUnderline = False
  frmMain.txtText.SelItalic = False
  lngLastLen& = Len(frmMain.txtText.Text)
  frmMain.txtText.SelStart = lngLastLen&
  frmMain.txtText.SelLength = 0
  frmMain.txtText.SelRTF = "   " & strText & vbNewLine
  frmMain.txtText.SelStart = lngLastLen&
  frmMain.txtText.SelLength = Len(frmMain.txtText.Text) - lngLastLen&
  frmMain.txtText.SelColor = strColor
  frmMain.txtText.SelFontSize = 8
  frmMain.txtText.SelBold = True
  frmMain.txtText.SelUnderline = False
  frmMain.txtText.SelItalic = False
  frmMain.txtText.SelHangingIndent = 1400
  frmMain.txtText.SelStart = Len(frmMain.txtText.Text)
  frmMain.txtText.SelLength = 0
End Sub

Sub AddOther(strData)
On Error Resume Next
  Dim lngLastLen As Long
  frmMain.txtText.SelStart = Len(frmMain.txtText.Text)
  frmMain.txtText.SelLength = 0
  frmMain.txtText.SelText = strData & vbNewLine
  frmMain.txtText.SelStart = Len(frmMain.txtText.Text) - (Len(strData))
  frmMain.txtText.SelLength = Len(strData) + 4
  frmMain.txtText.SelColor = &H800080
  frmMain.txtText.SelFontSize = 8
  frmMain.txtText.SelBold = True
  frmMain.txtText.SelUnderline = False
  frmMain.txtText.SelItalic = False
  frmMain.txtText.SelHangingIndent = 1400
  frmMain.txtText.SelStart = Len(frmMain.txtText.Text)
  frmMain.txtText.SelLength = 0
End Sub

Sub AddOther2(strData)
On Error Resume Next
  Dim lngLastLen As Long
  frmMain.txtText.SelStart = Len(frmMain.txtText.Text)
  frmMain.txtText.SelLength = 0
  frmMain.txtText.SelText = strData & vbNewLine
  frmMain.txtText.SelStart = Len(frmMain.txtText.Text) - (Len(strData))
  frmMain.txtText.SelLength = Len(strData) + 4
  frmMain.txtText.SelColor = &H800000
  frmMain.txtText.SelFontSize = 8
  frmMain.txtText.SelBold = True
  frmMain.txtText.SelUnderline = False
  frmMain.txtText.SelItalic = False
  frmMain.txtText.SelHangingIndent = 1400
  frmMain.txtText.SelStart = Len(frmMain.txtText.Text)
  frmMain.txtText.SelLength = 0
End Sub

Sub AddMOTD(strData As String)
On Error Resume Next
  Dim lngLastLen As Long
  frmMain.txtText.SelStart = Len(frmMain.txtText.Text)
  frmMain.txtText.SelLength = 0
  frmMain.txtText.SelText = strData & vbNewLine
  frmMain.txtText.SelStart = Len(frmMain.txtText.Text) - (Len(strData))
  frmMain.txtText.SelLength = Len(strData) + 4
  frmMain.txtText.SelColor = &H40C0&
  frmMain.txtText.SelFontSize = 8
  frmMain.txtText.SelBold = True
  frmMain.txtText.SelUnderline = False
  frmMain.txtText.SelItalic = False
  frmMain.txtText.SelHangingIndent = 1400
  frmMain.txtText.SelStart = Len(frmMain.txtText.Text)
  frmMain.txtText.SelLength = 0
End Sub


