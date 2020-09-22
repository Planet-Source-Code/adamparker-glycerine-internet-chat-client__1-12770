VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPref 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Preferences"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPref.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3840
   ScaleWidth      =   5505
   Visible         =   0   'False
   Begin MSComDlg.CommonDialog dlgColor 
      Left            =   1800
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&save"
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
      Left            =   4560
      TabIndex        =   5
      Top             =   3320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&cancel"
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
      Left            =   3720
      TabIndex        =   6
      Top             =   3320
      Width           =   855
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3570
      ItemData        =   "frmPref.frx":08CA
      Left            =   120
      List            =   "frmPref.frx":08D4
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&ok"
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
      Left            =   2880
      TabIndex        =   7
      Top             =   3320
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.CommandButton Command4 
         Caption         =   "remove"
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
         TabIndex        =   3
         Top             =   2280
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "add"
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
         Left            =   1680
         TabIndex        =   2
         Top             =   2280
         Width           =   855
      End
      Begin VB.ListBox lstBlock 
         Height          =   1620
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Block List:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   705
      End
   End
   Begin MSComDlg.CommonDialog cmdLoad 
      Left            =   3840
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Script Files (*.aos)|*.aos|"
   End
   Begin VB.Frame Frame3 
      Height          =   3015
      Left            =   1800
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   3615
      Begin VB.CheckBox Check1 
         Caption         =   "Play sounds"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   240
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmPref"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
End Sub

Private Sub Command10_Click()
dlgColor.ShowColor
Text1.ForeColor = dlgColor.Color
Text2.ForeColor = dlgColor.Color
End Sub

Private Sub Command12_Click()
dlgColor.ShowColor
List2.BackColor = dlgColor.Color
End Sub

Private Sub Command13_Click()
dlgColor.ShowColor
List2.ForeColor = dlgColor.Color
End Sub

Private Sub Command2_Click()
Call WriteToINI("Prefs", "Sound", Check1.Value, App.Path & "\prefs\Glycerine.ini")
Call SaveListBox(App.Path & "\prefs\ban.lst", lstBlock)
End Sub

Private Sub Command3_Click()
strInput = InputBox("Enter a name to block", "Glycerine Script Prompt")
lstBlock.AddItem (strInput)
End Sub

Private Sub Command4_Click()
lstBlock.RemoveItem (lstBlock.ListIndex)
End Sub

Private Sub Command5_Click()
cmdLoad.ShowOpen
txtStart.Clear
txtStart.AddItem (cmdLoad.FileName)
End Sub

Private Sub Command6_Click()
txtStart.Clear
End Sub

Private Sub Command7_Click()
Call WriteToINI("Prefs", "Sound", Check1.Value, App.Path & "\prefs\Glycerine.ini")
Call SaveListBox(App.Path & "\prefs\ban.lst", lstBlock)
Me.Hide
End Sub

Private Sub Command8_Click()
dlgColor.ShowColor
Command8.BackColor = dlgColor.Color
Command9.BackColor = dlgColor.Color
Command10.BackColor = dlgColor.Color
Command12.BackColor = dlgColor.Color
Command13.BackColor = dlgColor.Color
End Sub

Private Sub Command9_Click()
dlgColor.ShowColor
Text1.BackColor = dlgColor.Color
Text2.BackColor = dlgColor.Color
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2
On Error GoTo err:
If FileExists(App.Path & "\prefs\ban.lst") = False Then MsgBox ("" & App.Path & "\prefs\ban.lst"" was not found!"): Exit Sub
Call Loadlistbox(App.Path & "\prefs\ban.lst", lstBlock)
Check1.Value = GetFromINI("Prefs", "Sound", App.Path & "\prefs\Glycerine.ini")
Exit Sub
err:
Call ErrorBox(err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Command7_Click
Cancel = True
End Sub

Private Sub List1_Click()
    Select Case List1.Text
        Case "Block list"
            Frame1.Visible = True
            Frame3.Visible = False
        Case "Sounds"
            Frame1.Visible = False
            Frame3.Visible = True
    End Select
End Sub

