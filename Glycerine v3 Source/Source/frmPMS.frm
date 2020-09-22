VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPMS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "<!- scripting !->"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPMS.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5130
   ScaleWidth      =   5235
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   2160
      Top             =   5400
   End
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   5175
      Begin RichTextLib.RichTextBox txtText 
         Height          =   2055
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   3625
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmPMS.frx":08CA
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Send"
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
         Left            =   3960
         TabIndex        =   5
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox txtSend 
         Height          =   975
         Left            =   120
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   2400
         Width           =   4935
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   5175
      TabIndex        =   0
      Top             =   0
      Width           =   5235
      Begin VB.CommandButton Command2 
         Caption         =   "save as..."
         Height          =   375
         Left            =   4200
         TabIndex        =   8
         Top             =   540
         Width           =   975
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000C&
         X1              =   2990
         X2              =   2990
         Y1              =   -120
         Y2              =   1200
      End
      Begin VB.Label lblTime 
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3120
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblDate 
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3120
         TabIndex        =   6
         Top             =   0
         Width           =   2055
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   3000
         X2              =   3000
         Y1              =   960
         Y2              =   -120
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PM From:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   -60
         TabIndex        =   3
         Top             =   0
         Width           =   2940
      End
      Begin VB.Label lblwho 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "<!- script-!>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   0
         TabIndex        =   1
         Top             =   360
         Width           =   2820
         WordWrap        =   -1  'True
      End
   End
   Begin MSComDlg.CommonDialog dlgSave 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "All Files (*.*)|*.*"
   End
End
Attribute VB_Name = "frmPMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'There are some bug with the send
'It wont show what you send in the window
'I am in the process of fixing it and other things

Private Sub Command1_Click()
If Search_ListBox(lblwho, frmMain.lstUsers) = "-1" Then
Dim frm As New frmNotice
frm.Show
frm.txtError = "The user """ & lblwho & """ is not currently online"
Unload Me
Else
frmConnect.winsck.SendData ("pm-" & lblwho & ";" & frmConnect.txtName & ":     " & txtSend)
txtText = txtText & vbNewLine & frmConnect.txtName & ":     " & txtSend
txtSend = ""
txtSend.SetFocus
End If
End Sub

Private Sub Command2_Click()
dlgSave.ShowSave
Open dlgSave.FileName For Output As #1
Print #1, txtText
Close #1
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2
lblDate = "Date: " & Format(Date, "mm/dd/yy")
End Sub

Private Sub Timer1_Timer()
lblTime = "Time: " & Time
End Sub

Private Sub txtText_Change()
txtText.SelLength = Len(txtText)
End Sub

Private Sub txtText_GotFocus()
On Error Resume Next
txtSend.SetFocus
End Sub
