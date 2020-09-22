VERSION 5.00
Begin VB.Form frmTimeDate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert Time/Date"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
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
      Left            =   3000
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
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
      Left            =   3000
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.ListBox lstTime 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3765
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Formats"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmTimeDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmEdit.txtText.SelText = lstTime.Text
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2
Me.Icon = frmMain.Icon
lstTime.AddItem Format(Now, "long time")
lstTime.AddItem Format(Now, "short time")
lstTime.AddItem Format(Now, "medium time")
lstTime.AddItem Format(Now, "general date")
lstTime.AddItem Format(Now, "long date")
lstTime.AddItem Format(Now, "medium date")
lstTime.AddItem Format(Now, "short date")
lstTime.AddItem (Date)
lstTime.AddItem Format(Date, "dd - mm - yyyy")
lstTime.AddItem Format(Date, "dd-mm-yy")
lstTime.AddItem Format(Date, "dd/mm/yy")
lstTime.AddItem Format(Date, "dd/mm/yyyy")
lstTime.AddItem Format(Date, "dd/mm")
lstTime.AddItem Format(Date, "dd")
lstTime.AddItem Format(Date, "dd/yy")
lstTime.AddItem Format(Date, "mm/yy")
lstTime.AddItem Format(Time, "hh-mm-ss")
lstTime.AddItem Format(Time, "hh.mm.ss")
lstTime.AddItem Format(Time, "hh-mm")
End Sub
