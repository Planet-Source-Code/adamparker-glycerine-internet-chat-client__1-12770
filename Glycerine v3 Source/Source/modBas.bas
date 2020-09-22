Attribute VB_Name = "modBas"
Option Explicit
Global PicName, lb, rb, zoomactive As Boolean, BrushType, RepRed, RepGre, repBlu, Progress, NumSides, AtAngle, Rx, Ry, PolyX, PolyY, CopyX, CopyY
Global ImageArray(4, 1500, 1500) As Integer
Global X, Y As Integer
Global larrCol() As Long
Global Const CB_HEIGHT = 400
Global Const Pi = 3.14159265359
Public Const SRCCOPY = &HCC0020
Public Const SRCAND = &H8800C6
Public Const SRCINVERT = &H660046
Public Const SRCPAINT = &HEE0086
Public Const SRCERASE = &H4400328
Public Const WHITENESS = &HFF0062
Public Const BLACKNESS = &H42
Public Const ThisApp = "Stu Paint V2"
Public Const ThisKey = "Recent Files"
Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Declare Function CloseClipBoard Lib "user32" Alias "CloseClipboard" () As Long
Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As String) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Declare Function EmptyClipboard Lib "user32" () As Long
Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lppt As POINTAPI, ByVal nCount As Long) As Long
Declare Function TWAIN_AcquireToFilename Lib "EZTW32.DLL" (ByVal hwndApp%, ByVal bmpFileName$) As Integer
Declare Function TWAIN_SelectImageSource Lib "EZTW32.DLL" (ByVal hwndApp&) As Long
Declare Function TWAIN_AcquireToClipboard Lib "EZTW32.DLL" (ByVal hwndApp As Long, ByVal wPixTypes As Long) As Long
Declare Function TWAIN_IsAvailable Lib "EZTW32.DLL" () As Long
Declare Function TWAIN_EasyVersion Lib "EZTW32.DLL" () As Long
Public Type POINTAPI
    X As Long
    Y As Long
End Type

Sub Joined(Name)
On Error GoTo err:
Dim Res As String, strData As String
Res = GetFromINI("Prefs", "Join", App.Path & "\prefs\Glycerine.ini")
If Res = "1" Then
strData = ReadFile(App.Path & "\prefs\cust.dat")
strData = Replace(strData, "%n", Name)
strData = Replace(strData, "%t", Time)
strData = Replace(strData, "%d", Date)
strData = Replace(strData, "%v", App.Major & "." & App.Minor & "." & App.Revision)
strData = Replace(strData, "%r", Random(1, 1000))
strData = Replace(strData, "%s", frmConnect.cmbServer)
strData = Replace(strData, "%ip", frmConnect.winsck.LocalIP)
strData = Replace(strData, "%p", frmConnect.txtName)
frmConnect.winsck.SendData ("me- *" & frmConnect.txtName & ": " & strData)
End If
Exit Sub
err:
Call ErrorBox(err.Description)
End Sub

Sub OnCmd(Cmd)
On Error GoTo err:
Dim Res As String, strData As String
Res = GetFromINI("Prefs", "OnCmd", App.Path & "\prefs\Glycerine.ini")
If Res = "1" And Cmd = GetFromINI("Prefs", "OnCmdTxt", App.Path & "\prefs\Glycerine.ini") Then
strData = ReadFile(App.Path & "\prefs\cust.dat")
strData = Replace(strData, "%n", frmMain.lstUsers.Text)
strData = Replace(strData, "%t", Time)
strData = Replace(strData, "%d", Date)
strData = Replace(strData, "%v", App.Major & "." & App.Minor & "." & App.Revision)
strData = Replace(strData, "%r", Random(1, 1000))
strData = Replace(strData, "%s", frmConnect.cmbServer)
strData = Replace(strData, "%ip", frmConnect.winsck.LocalIP)
strData = Replace(strData, "%p", frmConnect.txtName)
frmConnect.winsck.SendData ("me- *" & frmConnect.txtName & ": " & strData)
End If
Exit Sub
err:
Call ErrorBox(err.Description)
End Sub

Function ReadFile(ByVal sFileName As String) As String
    Dim fhFile As Integer
    fhFile = FreeFile
    Open sFileName For Binary As #fhFile
    ReadFile = Input$(LOF(fhFile), fhFile)
    Close #fhFile
End Function

Function Search_ListBox(trig$, lst As ListBox) As Long
    Dim items As Long
    Dim n As Long
    items = lst.ListCount - 1
    For n = 0 To items Step 1
        If lst.List(n) = trig$ Then
            Search_ListBox = n
            Exit Function
        End If
    Next n
    Search_ListBox = -1
End Function

Sub ErrorBox(error)
Dim frm As New frmNotice
frm.Show
frm.txtError = "Notice: " & error
End Sub
