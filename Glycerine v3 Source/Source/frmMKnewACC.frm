VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMKnewACC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Make a new account"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4230
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMKnewACC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4620
   ScaleWidth      =   4230
   Begin VB.Frame Frame2 
      Height          =   4335
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   3975
      Begin VB.CommandButton Command3 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   16
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Next >>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "I agree to the terms."
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
         TabIndex        =   12
         Top             =   3480
         Width           =   2160
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   30
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3120
         TabIndex        =   18
         Top             =   2520
         Width           =   735
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00FFFFFF&
         X1              =   3010
         X2              =   3010
         Y1              =   1920
         Y2              =   3360
      End
      Begin VB.Line Line9 
         BorderColor     =   &H8000000C&
         X1              =   3000
         X2              =   3000
         Y1              =   1920
         Y2              =   3360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Porn No Warez"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3075
         TabIndex        =   17
         Top             =   2040
         Width           =   825
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line8 
         BorderColor     =   &H8000000C&
         X1              =   -120
         X2              =   3960
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00FFFFFF&
         X1              =   -120
         X2              =   3960
         Y1              =   855
         Y2              =   855
      End
      Begin VB.Line Line6 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   3960
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   3960
         Y1              =   3375
         Y2              =   3375
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   3960
         Y1              =   1935
         Y2              =   1935
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   3960
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label7 
         Caption         =   $"frmMKnewACC.frx":08CA
         Height          =   975
         Left            =   120
         TabIndex        =   15
         Top             =   2280
         Width           =   2895
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Terms:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   14
         Top             =   2040
         Width           =   660
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Please sign up for a account for Glycerine. You will not be able to connect until you do so."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   690
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   3705
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMKnewACC.frx":0995
         ForeColor       =   &H00000000&
         Height          =   780
         Left            =   135
         TabIndex        =   10
         Top             =   1005
         Width           =   3675
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5640
      Top             =   4680
   End
   Begin MSWinsockLib.Winsock winsck 
      Left            =   5640
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
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
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.ComboBox cmbServer 
         Height          =   315
         ItemData        =   "frmMKnewACC.frx":0A3D
         Left            =   1080
         List            =   "frmMKnewACC.frx":0A3F
         MousePointer    =   1  'Arrow
         TabIndex        =   7
         Top             =   2040
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Make account"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   5
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox txtPW 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1080
         MousePointer    =   1  'Arrow
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1080
         MaxLength       =   16
         MousePointer    =   1  'Arrow
         TabIndex        =   3
         Top             =   840
         Width           =   2655
      End
      Begin VB.Image Image1 
         Height          =   465
         Left            =   600
         Picture         =   "frmMKnewACC.frx":0A41
         Top             =   240
         Width           =   2565
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   3840
         Y1              =   3970
         Y2              =   3970
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   3840
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Server:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   540
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
         TabIndex        =   6
         Top             =   4080
         Width           =   3540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Password:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   1440
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Username:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmMKnewACC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
On Error GoTo err:
If txtName = "" Or txtPW = "" Or cmbServer = "" Then Call ErrorBox("Make sure all areas are filled in")
If winsck.State <> sckClosed Then Call winsck.Close
winsck.Connect cmbServer, "5623"
Exit Sub
err:
Call ErrorBox(err.Description)
End Sub

Private Sub Command2_Click()
If Check1.Value = "1" Then
Frame2.Visible = False
Frame1.Visible = True
Else
Call ErrorBox("You must agree to the terms!")
End If
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2
cmbServer.Text = frmConnect.cmbServer.Text
End Sub

Private Sub Timer1_Timer()
    If winsck.State = 0 Then
        lblStatus.Caption = "Not Connected."
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
    End If
    If winsck.State = 8 Then
        lblStatus.Caption = "Connection Terminated by Server."
    End If
    If winsck.State = 9 Then
        lblStatus.Caption = "Error!"
    End If

End Sub

Private Sub winsck_Connect()
On Error GoTo err:
Call CorrectName
winsck.SendData ("jrm-" & txtName & ":" & txtPW)
Exit Sub
err:
Call ErrorBox(err.Description)
End Sub

Private Sub winsck_DataArrival(ByVal bytesTotal As Long)
On Error GoTo err:
Dim Data As String
Call winsck.GetData(Data$, vbString)
frmDebug.txtDebug = Data$ & vbNewLine & vbNewLine & frmDebug.txtDebug
If Mid(Data$, 1, 6) = "accok-" Then
strData$ = Mid(Data$, 7)
Call CorrectName
If strData$ = txtName Then Call ErrorBox("Account """ & txtName & """ has been made!")
winsck.Close
frmConnect.txtName = txtName
frmConnect.txtPW = txtPW
Exit Sub
End If
If Mid(Data$, 1, 6) = "accno-" Then
strData$ = Mid(Data$, 7)
Call CorrectName
If strData$ = txtName Then Call ErrorBox("Account """ & txtName & """ is in use!")
winsck.Close
Exit Sub
End If
Exit Sub
err:
Call ErrorBox(err.Description)
End Sub

Private Sub winsck_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Call ErrorBox(Description)
End Sub
Private Sub CorrectName()
txtName = ReplaceString(txtName, "-", "_")
txtName = ReplaceString(txtName, ":", "_")
txtName = ReplaceString(txtName, ";", "_")
txtName = ReplaceString(txtName, " ", "_")
End Sub

