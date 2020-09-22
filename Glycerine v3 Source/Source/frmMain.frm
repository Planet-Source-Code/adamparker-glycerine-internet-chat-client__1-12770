VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Glycerine"
   ClientHeight    =   3900
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   8325
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkMode        =   1  'Source
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MouseIcon       =   "frmMain.frx":08CA
   ScaleHeight     =   3900
   ScaleWidth      =   8325
   Begin VB.PictureBox Picture1 
      Height          =   3135
      Left            =   240
      ScaleHeight     =   3075
      ScaleWidth      =   7755
      TabIndex        =   6
      Top             =   240
      Width           =   7815
      Begin VB.CommandButton Command4 
         Caption         =   "commands"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2205
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "load image"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   12
         Top             =   2205
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "send pm"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         Picture         =   "frmMain.frx":0A1C
         TabIndex        =   11
         Top             =   2205
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "change color"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   10
         Top             =   2205
         Width           =   1335
      End
      Begin VB.ListBox lstUsers 
         ForeColor       =   &H00000000&
         Height          =   1620
         ItemData        =   "frmMain.frx":2716
         Left            =   6120
         List            =   "frmMain.frx":2718
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   345
         Width           =   1575
      End
      Begin VB.TextBox txtData 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         MaxLength       =   200
         MousePointer    =   1  'Arrow
         TabIndex        =   8
         Top             =   2685
         Width           =   6375
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Send"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6600
         TabIndex        =   7
         Top             =   2685
         Width           =   1095
      End
      Begin RichTextLib.RichTextBox txtText 
         Height          =   1815
         Left            =   120
         TabIndex        =   14
         Top             =   165
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   3201
         _Version        =   393217
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         MousePointer    =   1
         TextRTF         =   $"frmMain.frx":271A
         MouseIcon       =   "frmMain.frx":27FE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   7920
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   6000
         X2              =   6000
         Y1              =   0
         Y2              =   2040
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   6015
         X2              =   6015
         Y1              =   0
         Y2              =   2040
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   7920
         Y1              =   2050
         Y2              =   2050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "People:"
         Height          =   195
         Left            =   6120
         TabIndex        =   16
         Top             =   120
         Width           =   540
      End
      Begin VB.Label lblRoom 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   6720
         TabIndex        =   15
         Top             =   120
         Width           =   975
      End
      Begin VB.Line Line5 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   7920
         Y1              =   2565
         Y2              =   2565
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   7920
         Y1              =   2580
         Y2              =   2580
      End
   End
   Begin VB.TextBox rtf3 
      Height          =   285
      Left            =   7680
      TabIndex        =   4
      Top             =   5520
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   8325
      TabIndex        =   1
      Top             =   3630
      Width           =   8325
      Begin VB.Line Line14 
         BorderColor     =   &H00FFFFFF&
         X1              =   10
         X2              =   10
         Y1              =   0
         Y2              =   240
      End
      Begin VB.Line Line13 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   240
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   8280
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line11 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   8280
         Y1              =   230
         Y2              =   230
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   8280
         Y1              =   10
         Y2              =   10
      End
      Begin VB.Line Line9 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   8280
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FFFFFF&
         X1              =   4455
         X2              =   4455
         Y1              =   360
         Y2              =   0
      End
      Begin VB.Line Line7 
         BorderColor     =   &H8000000C&
         X1              =   4440
         X2              =   4440
         Y1              =   360
         Y2              =   0
      End
      Begin VB.Label lblNick 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "- nick -"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4560
         TabIndex        =   3
         Top             =   0
         Width           =   3615
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "- stats -"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   75
         TabIndex        =   2
         Top             =   20
         Width           =   4185
      End
      Begin VB.Line Line16 
         BorderColor     =   &H00FFFFFF&
         X1              =   8270
         X2              =   8270
         Y1              =   0
         Y2              =   240
      End
      Begin VB.Line Line15 
         BorderColor     =   &H8000000C&
         X1              =   8280
         X2              =   8280
         Y1              =   0
         Y2              =   240
      End
   End
   Begin MSComDlg.CommonDialog dlgColor 
      Left            =   8520
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Text Files (*.txt)|*.txt"
   End
   Begin VB.TextBox txtLoad 
      CausesValidation=   0   'False
      Height          =   495
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   5640
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog cmdLoad 
      Left            =   8520
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "All Files (*.*)|*.*|"
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   8520
      Top             =   960
   End
   Begin MSWinsockLib.Winsock winsck 
      Left            =   8520
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Label lblColor 
      Height          =   135
      Left            =   960
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image imgSkin 
      Height          =   3615
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim htxt As String
Dim word As String
Const SW_SHOWNORMAL = 1
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const EM_CHARFROMPOS& = &HD7
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Sub Command1_Click()
frmMSG.Show
frmMSG.txtWho = lstUsers.Text
End Sub

Private Sub Command2_Click()
On Error Resume Next
cmdLoad.ShowOpen
imgSkin = LoadPicture(cmdLoad.FileName)
End Sub

Private Sub Command3_Click()
Dim strToSend As String
If txtData = "" Then Exit Sub
On Error GoTo err:
If LCase(Mid(txtData, 1, 5)) = LCase("/back") Then
strToSend$ = Mid(txtData, 6)
TimeOut 0.2
frmConnect.winsck.SendData ("me-* " & frmConnect.txtName & "  is back.")
txtData = ""
Exit Sub
End If
If LCase(Mid(txtData, 1, 7)) = LCase("/nobeer") Then
strToSend$ = Mid(txtData, 8)
TimeOut 0.2
frmConnect.winsck.SendData ("me-* " & frmConnect.txtName & "  opens a can of 100% Non-Alchoholic Canadian beer.")
txtData = ""
Exit Sub
End If
If LCase(Mid(txtData, 1, 5)) = LCase("/beer") Then
strToSend$ = Mid(txtData, 6)
TimeOut 0.2
frmConnect.winsck.SendData ("me-* " & frmConnect.txtName & " opens a can of beer.")
txtData = ""
Exit Sub
End If
If LCase(Mid(txtData, 1, 5)) = LCase("/lmao") Then
strToSend$ = Mid(txtData, 6)
TimeOut 0.2
frmConnect.winsck.SendData ("me-* " & frmConnect.txtName & " is laughing his ass off.")
txtData = ""
Exit Sub
End If
If LCase(Mid(txtData, 1, 5)) = LCase("/rofl") Then
strToSend$ = Mid(txtData, 6)
TimeOut 0.2
frmConnect.winsck.SendData ("me-* " & frmConnect.txtName & " is rolling on the floor laughing.")
txtData = ""
Exit Sub
End If
If LCase(Mid(txtData, 1, 4)) = LCase("/brb") Then
strToSend$ = Mid(txtData, 5)
TimeOut 0.2
frmConnect.winsck.SendData ("me-* " & frmConnect.txtName & " will be right back.")
txtData = ""
Exit Sub
End If
If LCase(Mid(txtData, 1, 4)) = LCase("/lol") Then
strToSend$ = Mid(txtData, 5)
TimeOut 0.2
frmConnect.winsck.SendData ("me-* " & frmConnect.txtName & " is laughing out loud.")
txtData = ""
Exit Sub
End If
If LCase(Mid(txtData, 1, 4)) = LCase("/keg") Then
strToSend$ = Mid(txtData, 5)
TimeOut 0.2
frmConnect.winsck.SendData ("me-* " & frmConnect.txtName & " taps a keg and starts drinking.")
txtData = ""
Exit Sub
End If
If LCase(Mid(txtData, 1, 4)) = LCase("/me ") Then
strToSend$ = Mid(txtData, 5)
TimeOut 0.2
frmConnect.winsck.SendData ("me-* " & frmConnect.txtName & " " & strToSend$)
txtData = ""
Exit Sub
End If
If LCase(Mid(txtData, 1, 1)) = LCase("/") Then
strToSend$ = Mid(txtData, 2)
TimeOut 0.2
Call OnCmd(strToSend$)
txtData = ""
Exit Sub
End If
TimeOut 0.2
frmConnect.winsck.SendData ("msg-" & frmConnect.txtName & ":    color=" & lblColor & "~" & txtData)
txtData = ""
txtText.SelLength = Len(txtText)
Exit Sub
err:
Call ErrorBox(err.Description)
End Sub

Private Sub Command4_Click()
frmCommands.Show
End Sub

Private Sub Command5_Click()
frmColor2.Show
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2
lblColor = GetSetting("Glycerine", "Color", "Code")
If lblColor = "" Then lblColor = "0"
txtData.ForeColor = lblColor
Call AddOther2("Welcome To Glycerine Online Chat!")
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
frmConnect.winsck.SendData "exit-" & frmConnect.txtName
DoEvents
frmConnect.winsck.Close
frmConnect.Show
Exit Sub
End Sub

Private Sub lstUsers_DblClick()
frmMSG.Show
frmMSG.txtWho = lstUsers.Text
End Sub

Private Sub lstUsers_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmMenu.mnuOptPM.Caption = "Send PM to " & lstUsers.Text
frmMenu.mnuOptSlap.Caption = "Slap " & lstUsers.Text
frmMenu.mnuLOL.Caption = "LOL at " & lstUsers.Text
frmMenu.mnuScare.Caption = "Scare " & lstUsers.Text
frmMenu.mnuPunch.Caption = "Punch " & lstUsers.Text
If Button = "2" Then PopupMenu frmMenu.mnuOpt
End Sub

Private Sub Timer1_Timer()
    If frmConnect.winsck.State = 0 Then
        lblStatus.Caption = "Not Connected."
        Unload Me
    End If
    If frmConnect.winsck.State = 1 Then
        lblStatus.Caption = "Socket Open."
    End If
    If frmConnect.winsck.State = 3 Then
        lblStatus.Caption = "Connection Pending..."
    End If
    If frmConnect.winsck.State = 4 Then
        lblStatus.Caption = "Resolving Host..."
    End If
    If frmConnect.winsck.State = 5 Then
        lblStatus.Caption = "Host Resolved."
    End If
    If frmConnect.winsck.State = 6 Then
        lblStatus.Caption = "Connecting..."
    End If
    If frmConnect.winsck.State = 7 Then
        lblStatus.Caption = "Connected."
    End If
    If frmConnect.winsck.State = 8 Then
        lblStatus.Caption = "Connection Terminated by Server."
        Unload Me
        frmConnect.Show
    End If
    If frmConnect.winsck.State = 9 Then
        lblStatus.Caption = "Error!"
    End If
End Sub

Private Sub txtData_KeyPress(KeyAscii As Integer)
If KeyAscii = "13" Then
Command3_Click
End If
End Sub

Private Sub txtText_Change()

End Sub

Private Sub txtText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = "2" Then PopupMenu frmMenu.mnuEdit
End Sub
