VERSION 5.00
Begin VB.Form frmLoad 
   BorderStyle     =   0  'None
   Caption         =   "spgc is loading"
   ClientHeight    =   2220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4860
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2220
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6120
      Top             =   1080
   End
   Begin VB.TextBox txtLoad 
      Height          =   855
      Left            =   5760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "loading..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   4725
   End
   Begin VB.Label lblVer 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Version:"
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
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   4695
   End
   Begin VB.Image logo 
      Height          =   2205
      Left            =   0
      Picture         =   "frmLoad.frx":0000
      Top             =   0
      Width           =   4875
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2
lblVer = "Version: " & App.Major & "." & App.Minor & "." & App.Revision
Me.Show
FormOnTop Me
Me.Width = logo.Width
Me.Height = logo.Height
If App.PrevInstance = True Then MsgBox "Glycerine is already open! If you closed the chat try using ctrl+alt+del to see if it is still running.": End
frmPref.Show
frmPref.Hide
If Not FileExists(App.Path & "\Glycerine.dat") = True Then MsgBox "Error! No Glycerine.dat file! Unloading!": End
Call LoadText(txtLoad, App.Path & "\Glycerine.dat")
TimeOut 1
mdiMain.Show
DoScript txtLoad
Unload Me
End Sub
