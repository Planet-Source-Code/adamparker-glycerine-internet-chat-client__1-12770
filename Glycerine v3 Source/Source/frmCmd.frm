VERSION 5.00
Begin VB.Form frmCmd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Action"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6900
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCmd.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4395
   ScaleWidth      =   6900
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
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
      Left            =   5400
      TabIndex        =   18
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox txtText 
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   4575
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "appears as: * NICK: -custom-command-"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   4185
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   4575
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "%p = your name"
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
         Left            =   2400
         TabIndex        =   11
         Top             =   960
         Width           =   1440
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "%ip = your ip"
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
         TabIndex        =   10
         Top             =   960
         Width           =   1155
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "%s = server"
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
         Left            =   2400
         TabIndex        =   7
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "%r = random number"
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
         Left            =   2400
         TabIndex        =   6
         Top             =   480
         Width           =   1860
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "%v = current version"
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
         Left            =   2400
         TabIndex        =   5
         Top             =   240
         Width           =   1830
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   2295
         X2              =   2295
         Y1              =   120
         Y2              =   1320
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   2280
         X2              =   2280
         Y1              =   120
         Y2              =   1320
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "%d = current date"
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
         TabIndex        =   4
         Top             =   720
         Width           =   1590
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "%t = current time"
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
         Top             =   480
         Width           =   1560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "%n = selected name"
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
         TabIndex        =   2
         Top             =   240
         Width           =   1770
      End
   End
   Begin VB.Frame Frame4 
      Height          =   3735
      Left            =   4800
      TabIndex        =   12
      Top             =   0
      Width           =   1935
      Begin VB.CheckBox Check1 
         Caption         =   "Auto say"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2640
         Width           =   1455
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Say on command"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox txtOnCmd 
         Height          =   285
         Left            =   240
         TabIndex        =   14
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   2640
         Y1              =   2530
         Y2              =   2530
      End
      Begin VB.Line Line5 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   2640
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "/"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   3240
         Width           =   60
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Other commands:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmCmd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Open App.Path & "\prefs\cust.dat" For Output As #1
Print #1, txtText
Close #1
Call WriteToINI("Prefs", "Join", Check1.Value, App.Path & "\prefs\Glycerine.ini")
Call WriteToINI("Prefs", "OnCmd", Check2.Value, App.Path & "\prefs\Glycerine.ini")
Call WriteToINI("Prefs", "OnCmdTxt", txtOnCmd, App.Path & "\prefs\Glycerine.ini")
Unload Me
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2
On Error Resume Next
Me.Icon = frmMain.Icon
txtText = ReadFile(App.Path & "\prefs\cust.dat")
Check1.Value = GetFromINI("Prefs", "Join", App.Path & "\prefs\Glycerine.ini")
Check2.Value = GetFromINI("Prefs", "OnCmd", App.Path & "\prefs\Glycerine.ini")
txtOnCmd = GetFromINI("Prefs", "OnCmdTxt", App.Path & "\prefs\Glycerine.ini")
End Sub

