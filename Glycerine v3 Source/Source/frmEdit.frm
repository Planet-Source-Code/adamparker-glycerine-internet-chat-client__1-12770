VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Glycerine Edit Pad"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9915
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEdit.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   9915
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H80000000&
      Height          =   5835
      Left            =   0
      ScaleHeight     =   5775
      ScaleWidth      =   1995
      TabIndex        =   6
      Top             =   0
      Width           =   2055
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   5535
         Left            =   120
         ScaleHeight     =   5535
         ScaleWidth      =   1815
         TabIndex        =   8
         Top             =   120
         Width           =   1815
         Begin VB.Line Line2 
            BorderColor     =   &H8000000C&
            X1              =   120
            X2              =   2280
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            X1              =   120
            X2              =   2280
            Y1              =   1215
            Y2              =   1215
         End
         Begin VB.Line Line5 
            BorderColor     =   &H8000000C&
            X1              =   120
            X2              =   2280
            Y1              =   3000
            Y2              =   3000
         End
         Begin VB.Line Line6 
            BorderColor     =   &H00FFFFFF&
            X1              =   120
            X2              =   2280
            Y1              =   3015
            Y2              =   3015
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "&New"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            MouseIcon       =   "frmEdit.frx":08CA
            MousePointer    =   99  'Custom
            TabIndex        =   15
            Top             =   120
            Width           =   1575
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "&Save"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            MouseIcon       =   "frmEdit.frx":0BD4
            MousePointer    =   99  'Custom
            TabIndex        =   14
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "&Open"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            MouseIcon       =   "frmEdit.frx":0EDE
            MousePointer    =   99  'Custom
            TabIndex        =   13
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "&Cut"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            MouseIcon       =   "frmEdit.frx":11E8
            MousePointer    =   99  'Custom
            TabIndex        =   12
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "&Copy"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            MouseIcon       =   "frmEdit.frx":14F2
            MousePointer    =   99  'Custom
            TabIndex        =   11
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "&Paste"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            MouseIcon       =   "frmEdit.frx":17FC
            MousePointer    =   99  'Custom
            TabIndex        =   10
            Top             =   2040
            Width           =   1575
         End
         Begin VB.Line Line3 
            BorderColor     =   &H8000000C&
            X1              =   120
            X2              =   2160
            Y1              =   2400
            Y2              =   2400
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00FFFFFF&
            X1              =   120
            X2              =   2160
            Y1              =   2415
            Y2              =   2415
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "&Insert Time/Date"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            MouseIcon       =   "frmEdit.frx":1B06
            MousePointer    =   99  'Custom
            TabIndex        =   9
            Top             =   2520
            Width           =   1575
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5415
      Left            =   2160
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      Begin VB.TextBox txtText 
         Height          =   5055
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   240
         Width           =   7455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2700
         TabIndex        =   5
         Top             =   510
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1200
         TabIndex        =   4
         Top             =   540
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label Label3 
         Caption         =   "Current Word"
         Height          =   255
         Left            =   2880
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Special Chars"
         Height          =   255
         Left            =   1350
         TabIndex        =   2
         Top             =   390
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin MSComDlg.CommonDialog dlgColor 
      Left            =   2880
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "All Files (*.*)|*.*"
   End
   Begin MSComDlg.CommonDialog dlgImg 
      Left            =   2400
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "All Files (*.*)|*.*"
   End
   Begin MSComctlLib.ImageList imglst 
      Left            =   2280
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdit.frx":1E10
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdit.frx":1F24
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdit.frx":2038
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdit.frx":214C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdit.frx":2260
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdit.frx":2374
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdit.frx":2488
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblLine 
      BackStyle       =   0  'Transparent
      Caption         =   "Line Count: 1"
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
      Left            =   2280
      TabIndex        =   7
      Top             =   5530
      Width           =   7575
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H80000018&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   2160
      Top             =   5520
      Width           =   7695
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Label10_Click()
txtText.SelText = Clipboard.GetText
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.ForeColor = &HFF&
End Sub

Private Sub Label10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.ForeColor = &HFF0000
End Sub

Private Sub Label11_Click()
frmTimeDate.Show
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label11.ForeColor = &HFF&
End Sub

Private Sub Label11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label11.ForeColor = &HFF0000
End Sub

Private Sub Label5_Click()
            If MsgBox("Are you sure? All unsaved changes will be lost!", vbYesNo) = vbYes Then txtText = ""
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.ForeColor = &HFF&
End Sub

Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.ForeColor = &HFF0000
End Sub

Private Sub Label6_Click()
            dlgColor.ShowSave
            Call SaveText(txtText, dlgColor.FileName)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = &HFF&
End Sub

Private Sub Label6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = &HFF0000
End Sub

Private Sub Label7_Click()
            dlgColor.ShowOpen
            Call LoadText(txtText, dlgColor.FileName)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.ForeColor = &HFF&
End Sub

Private Sub Label7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.ForeColor = &HFF0000
End Sub

Private Sub Label8_Click()
Clipboard.SetText txtText.SelText
txtText.SelText = ""
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.ForeColor = &HFF&
End Sub

Private Sub Label8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.ForeColor = &HFF0000
End Sub

Private Sub Label9_Click()
Clipboard.SetText txtText.SelText
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.ForeColor = &HFF&
End Sub

Private Sub Label9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.ForeColor = &HFF0000
End Sub

Private Sub txtText_Change()
lblLine = "Line Count: " & LineCount(txtText)
End Sub
