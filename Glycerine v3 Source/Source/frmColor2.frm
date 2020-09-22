VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmColor2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pick a color..."
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2580
   Icon            =   "frmColor2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   2580
   Begin VB.Frame Frame1 
      Caption         =   "color chart:"
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      Begin VB.TextBox txtColor 
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
         Left            =   120
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   4
         Text            =   "0"
         Top             =   3120
         Width           =   2295
      End
      Begin VB.CommandButton Command5 
         Caption         =   "ok"
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
         TabIndex        =   3
         Top             =   3480
         Width           =   855
      End
      Begin VB.PictureBox Picture3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Height          =   2295
         Left            =   120
         ScaleHeight     =   2235
         ScaleWidth      =   2235
         TabIndex        =   1
         ToolTipText     =   "Click here to pick a color"
         Top             =   240
         Width           =   2295
         Begin VB.CommandButton Command1 
            Caption         =   "click to pick color"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   0
            TabIndex        =   2
            Top             =   1920
            Width           =   2230
         End
      End
      Begin MSComDlg.CommonDialog dlgColor 
         Left            =   4680
         Top             =   3600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   2520
         Y1              =   2770
         Y2              =   2770
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   2520
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "color code:"
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
         TabIndex        =   5
         Top             =   2880
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmColor2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim initVals(4) As Long

Private Sub Command1_Click()
Picture3_Click
End Sub

Private Sub Command5_Click()
frmMain.lblColor = txtColor.Text
frmMain.txtData.ForeColor = txtColor.Text
Call SaveSetting("Glycerine", "Color", "Code", txtColor.Text)
Unload Me
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2
Dim intBuffer As Integer
Dim Acolor As Long
        Picture3.BackColor = frmMain.lblColor.Caption
        Acolor = frmMain.lblColor.Caption
        initVals(3) = Acolor
        initVals(0) = initVals(3) And 255
        initVals(1) = (initVals(3) And 65280) \ 256&
        initVals(2) = (initVals(3) And 16711680) \ 65535
        reFreshAll initVals(0), initVals(1), initVals(2)
End Sub

Private Sub Picture3_Click()
    Dim Acolor As Long
    With dlgColor
        .FLAGS = cdlCCRGBInit Or cdlCCFullOpen
        .CancelError = False
        .Color = Picture3.BackColor
        .ShowColor
        Acolor = .Color
    End With
    If Acolor <> Picture3.BackColor Then
        initVals(3) = Acolor
        initVals(0) = initVals(3) And 255
        initVals(1) = (initVals(3) And 65280) \ 256&
        initVals(2) = (initVals(3) And 16711680) \ 65535
        reFreshAll initVals(0), initVals(1), initVals(2)
    End If

End Sub

Private Sub reFreshAll(Optional Red As Long, _
        Optional Green As Long, _
        Optional Blue As Long, _
        Optional cValue As Variant)
    Dim Color As Long
    If Not IsMissing(cValue) Then
        Color = cValue
        Red = Color And 255
        Green = (Color And 65280) \ 256&
        Blue = (Color And 16711680) \ 65535
    Else
        Color = Red + Green * 256 + Blue * 256 * 256
    End If
    txtColor.Text = CStr(Color)
    initVals(0) = Red
    initVals(1) = Green
    initVals(2) = Blue
    initVals(3) = Color
    Picture3.BackColor = Color
End Sub
