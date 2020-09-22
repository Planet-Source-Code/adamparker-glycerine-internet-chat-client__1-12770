VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConnect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Glycerine - Connect"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4620
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConnect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3795
   ScaleWidth      =   4620
   Begin VB.CommandButton Command3 
      Caption         =   "signup!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      Picture         =   "frmConnect.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "sign on!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      Picture         =   "frmConnect.frx":170C
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5640
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSComctlLib.ImageList imglst 
      Left            =   4800
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConnect.frx":254E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConnect.frx":33A2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3795
      Left            =   0
      ScaleHeight     =   3795
      ScaleWidth      =   1125
      TabIndex        =   10
      Top             =   0
      Width           =   1125
      Begin VB.Image Image2 
         Height          =   975
         Left            =   75
         Picture         =   "frmConnect.frx":41F6
         Top             =   75
         Width           =   975
      End
      Begin VB.Label lblVer 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   915
         Left            =   0
         TabIndex        =   15
         Top             =   1080
         Width           =   1125
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Glycerine Connect Window"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   585
         Left            =   120
         TabIndex        =   11
         Top             =   3120
         Width           =   915
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   1200
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin VB.CheckBox Check1 
         Caption         =   "&Save Password"
         Height          =   255
         Left            =   1800
         TabIndex        =   9
         Top             =   2160
         Width           =   1455
      End
      Begin VB.ComboBox cmbServer 
         Height          =   315
         ItemData        =   "frmConnect.frx":4CBC
         Left            =   960
         List            =   "frmConnect.frx":4CBE
         MousePointer    =   1  'Arrow
         TabIndex        =   8
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox txtPW 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         MousePointer    =   1  'Arrow
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   960
         MaxLength       =   16
         MousePointer    =   1  'Arrow
         TabIndex        =   6
         Top             =   360
         Width           =   2295
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   540
         Left            =   1200
         TabIndex        =   14
         Top             =   2640
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   953
         ButtonWidth     =   1667
         ButtonHeight    =   953
         Style           =   1
         ImageList       =   "imglst"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Sign On  "
               Key             =   "Sign On"
               Object.ToolTipText     =   "Sign On"
               Object.Tag             =   "Sign On"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Join Today!"
               Key             =   "Sign Up"
               Object.ToolTipText     =   "Sign Up"
               Object.Tag             =   "Sign Up"
               ImageIndex      =   1
            EndProperty
         EndProperty
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   6720
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   6720
         Y1              =   3375
         Y2              =   3375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Password:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   750
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         Caption         =   "- status -"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   3480
         Width           =   3060
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Server:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   1560
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Username:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   780
      End
   End
   Begin VB.Timer tmrStat 
      Interval        =   1
      Left            =   5640
      Top             =   5040
   End
   Begin MSWinsockLib.Winsock winsck 
      Left            =   3960
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label idle 
      AutoSize        =   -1  'True
      Caption         =   "off"
      Height          =   195
      Left            =   3720
      TabIndex        =   4
      Top             =   4800
      Visible         =   0   'False
      Width           =   210
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private strUsers() As String
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

Private Sub cmbServer_KeyPress(KeyAscii As Integer)
If KeyAscii = "13" Then Command1_Click
End Sub

Private Sub Command1_Click()
On Error GoTo err:
If Check1.Value = "1" Then
Call SaveSetting("Glycerine", "Login", "Name", txtName)
Call SaveSetting("Glycerine", "Login", "Password", txtPW)
Call SaveSetting("Glycerine", "Login", "Save", "1")
Else
Call SaveSetting("Glycerine", "Login", "Save", "0")
End If
If txtName = "" Or txtPW = "" Or cmbServer = "" Then MsgBox ("Make sure all areas are filled in"): Exit Sub
txtPW = ReplaceString(txtPW, "_", "")
If winsck.State <> sckClosed Then Call winsck.Close
winsck.Connect cmbServer, "5622"
Exit Sub
err:
Call ErrorBox(err.Description)
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2
If GetSetting("Glycerine", "Login", "Save") = "1" Then
txtName = GetSetting("Glycerine", "Login", "Name")
txtPW = GetSetting("Glycerine", "Login", "Password")
Check1.Value = "1"
End If
lblVer = App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.WindowState = 1
Cancel = True
End Sub

Private Sub tmrStat_Timer()
    If winsck.State = 0 Then
        lblStatus.Caption = "Not Connected."
        mdiMain.mnuConnect.Enabled = True
    End If
    If winsck.State = 1 Then
        lblStatus.Caption = "Socket Open."
    End If
    If winsck.State = 3 Then
        lblStatus.Caption = "Connection Pending..."
    End If
    If winsck.State = 4 Then
        lblStatus.Caption = "Resolving Host..."
    End If
    If winsck.State = 5 Then
        lblStatus.Caption = "Host Resolved."
    End If
    If winsck.State = 6 Then
        lblStatus.Caption = "Connecting..."
    End If
    If winsck.State = 7 Then
        lblStatus.Caption = "Connected."
        mdiMain.mnuConnect.Enabled = False
        Me.Visible = False
    End If
    If winsck.State = 8 Then
        lblStatus.Caption = "Connection Terminated by Server."
        mdiMain.mnuConnect.Enabled = True
    End If
    If winsck.State = 9 Then
        lblStatus.Caption = "Error!"
        mdiMain.mnuConnect.Enabled = True
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Sign On"
            Command1_Click
        Case "Sign Up"
            frmMKnewACC.Show
    End Select
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = "13" Then Command1_Click
End Sub

Private Sub txtPW_KeyPress(KeyAscii As Integer)
If KeyAscii = "13" Then Command1_Click
End Sub

Private Sub winsck_Connect()
On Error GoTo err:
strLogin = EncryptA(txtName & ":" & txtPW)
winsck.SendData ("req-" & strLogin)
frmMain.lblNick = txtName
Exit Sub
err:
Call ErrorBox(err.Description)
End Sub

Private Sub winsck_DataArrival(ByVal bytesTotal As Long)
On Error GoTo err:
Dim Data As String
Call winsck.GetData(Data$, vbString)
Call StepThrough(Data$)
Exit Sub
err:
Call ErrorBox(err.Description)
End Sub

Private Sub winsck_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Call ErrorBox(err.Description)
End Sub

Sub StepThrough(Data As String)
On Error GoTo err:
If Not Mid(Data$, 1, 3) = "pm-" Then frmDebug.txtDebug = Data$ & vbNewLine & frmDebug.txtDebug

If Mid(Data$, 1, 3) = "me-" Then
strData$ = Mid(Data$, 4)
Call AddOther(strData$)
End If

If Mid(Data$, 1, 4) = "msg-" Then
strData$ = Mid(Data$, 5)
strName$ = Left$(strData, InStr(strData, ":") - 1)
If Search_ListBox(strName$, frmPref.lstBlock) = "-1" Then
If strName$ = txtName Then
If frmPref.Check1.Value = 1 Then Call Playwav(App.Path & "\sounds\send.wav")
Call AddChat(strData$)
Else
If frmPref.Check1.Value = 1 Then Call Playwav(App.Path & "\sounds\get.wav")
Call AddChat2(strData$)
End If
End If
Else
End If

If Mid(Data$, 1, 4) = "adm-" Then
If frmPref.Check1.Value = 1 Then Call Playwav(App.Path & "\sounds\admin.wav")
strData$ = Mid(Data$, 5)
Dim frmm As New frmAdmin
frmm.Show
frmm.Text1 = strData$
End If

If Mid(Data$, 1, 3) = "pm-" Then
strData = Mid(Data$, 4)
strUser = Left$(strData, InStr(strData, ";") - 1)
strData = Right(strData, Len(strData) - InStr(strData, ";"))
strUser2 = Left$(strData, InStr(strData, ":") - 1)
strData2 = Right(strData, Len(strData) - InStr(strData, ":"))
If Not LCase(strUser) = LCase(txtName) Then Exit Sub
For Each frm In Forms
If LCase(frm.Tag) = LCase(strUser2) Then
frm.txtText.Text = frm.txtText.Text & vbNewLine & strUser2 & ":" & strData2
FoundChat = 1
End If
Next frm
If FoundChat = 0 Then
Dim myForm As New frmPMS
Load myForm
myForm.Caption = strUser2 & " - Private Message"
myForm.Tag = strUser2
myForm.Visible = True
myForm.txtText.Text = myForm.txtText.Text & vbNewLine & strUser2 & ":" & strData2
myForm.lblwho = strUser2
End If
Exit Sub

If strUser = frmConnect.txtName Then
Call AddOther("Private Message - " & strData)
If idle = "on" Then
frmPMLog.Text1 = "[" & Time & " / " & Date & "] " & strData$ & vbNewLine & vbNewLine & frmPMLog.Text1
frmPMLog.Show
End If
End If
End If

If Mid(Data$, 1, 6) = "users-" Then
strData = Mid(Data$, 7)
strNames = strData
Do
a = InStr(strNames, "-")
If a = 0 Then Exit Do
w = Left(strNames, a - 1)
b = Left(strNames, a)
c = Len(strNames)
d = c - a
z = Right(strNames, d)
strNames = z
If Not w = "" Or w = " " Then frmMain.lstUsers.AddItem (w)
Loop Until a = 0
frmMain.lblRoom = frmMain.lstUsers.ListCount
End If

If Mid(Data$, 1, 5) = "exit-" Then
If frmPref.Check1.Value = 1 Then Call Playwav(App.Path & "\sounds\on.wav")
strExit$ = Mid(Data, 6)
If strExit$ = txtName Then winsck.Close: frmConnect.Show: Unload frmMain
Call AddOther(strExit$ & " has left SPG Chat [" & Time & "]")
For n = 0 To frmMain.lstUsers.ListCount - 1 Step 1
If frmMain.lstUsers.List(n) = strExit$ Then
frmMain.lstUsers.RemoveItem (n)
frmMain.lblRoom = frmMain.lstUsers.ListCount
Exit Sub
End If
Next n
End If

If Mid(Data$, 1, 5) = "join-" Then
If frmPref.Check1.Value = 1 Then Call Playwav(App.Path & "\sounds\on.wav")
strAdd = Mid(Data$, 6)
If strAdd = txtName Then Exit Sub
frmMain.lstUsers.AddItem (strAdd)
frmMain.lblRoom = frmMain.lstUsers.ListCount
Call AddOther(strAdd & " has joined SPG Chat [" & Time & "]")
If Not strAdd = frmConnect.txtName Then Call Joined(strAdd)
End If

If Mid(Data$, 1, 5) = "motd-" Then
strMOTD$ = Mid(Data$, 6)
frmMOTD.Text1 = strMOTD$
frmMOTD.Show
End If

If Mid(Data$, 1, 3) = "ok-" Then
strUser = Mid(Data$, 4)
If strUser = txtName Then
Me.Hide
frmMain.Show
End If
End If

If Mid(Data$, 1, 3) = "wb-" Then
strSite = Mid(Data$, 4)
If MsgBox("Do you want to visit the website """ & strSite & """?", vbYesNo) = vbYes Then
Dim Popup As New frmPopUp
Popup.Show
Popup.WebBrowser1.Navigate (strSite)
End If
End If

If Mid(Data$, 1, 6) = "inuse-" Then
strUser = Mid(Data$, 7)
If strUser = txtName Then
winsck.Close
Unload frmMain
Call ErrorBox(err.Description)
End If
End If

Exit Sub
err:
Call ErrorBox(err.Description)
End Sub

