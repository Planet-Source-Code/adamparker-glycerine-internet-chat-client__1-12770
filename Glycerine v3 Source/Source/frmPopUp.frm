VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPopUp 
   Caption         =   "Glycerine Browser"
   ClientHeight    =   4905
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   8055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPopUp.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   4905
   ScaleWidth      =   8055
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   2640
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   12632256
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopUp.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopUp.frx":0BAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopUp.frx":0E8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopUp.frx":1170
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopUp.frx":1452
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopUp.frx":1734
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopUp.frx":1A16
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopUp.frx":1B36
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   8055
      TabIndex        =   3
      Top             =   480
      Width           =   8055
      Begin VB.CommandButton Command5 
         Caption         =   "go"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   305
         Left            =   4320
         TabIndex        =   5
         Top             =   15
         Width           =   495
      End
      Begin VB.ComboBox cmbAddress 
         Height          =   315
         Left            =   840
         MousePointer    =   1  'Arrow
         TabIndex        =   4
         Top             =   0
         Width           =   3375
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   4920
         TabIndex        =   7
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   8040
         Y1              =   390
         Y2              =   390
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   8040
         Y1              =   380
         Y2              =   380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ad&dress:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   50
         Width           =   645
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   847
      ButtonWidth     =   820
      ButtonHeight    =   794
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "back"
            Object.ToolTipText     =   "back"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "forward"
            Object.ToolTipText     =   "forward"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "refresh"
            Object.ToolTipText     =   "refresh"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "stop"
            Object.ToolTipText     =   "stop"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stats 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   1
      Top             =   4635
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11139
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "11:32 PM"
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   8055
      ExtentX         =   14208
      ExtentY         =   6376
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin MSComDlg.CommonDialog cmdLoad 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "HTML Files (*.html)|*.html"
   End
End
Attribute VB_Name = "frmPopUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbAddress_Change()
Call SaveListBox(App.Path & "\prefs\amcr.dat", cmbAddress)
End Sub

Private Sub cmbAddress_KeyPress(KeyAscii As Integer)
If KeyAscii = "13" Then Command5_Click
End Sub

Private Sub Command5_Click()
WebBrowser1.Navigate (cmbAddress.Text)
cmbAddress.AddItem (cmbAddress.Text)
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2
Call Loadlistbox(App.Path & "\prefs\amcr.dat", cmbAddress)
End Sub

Private Sub Form_Resize()
On Error Resume Next
WebBrowser1.Width = Me.Width - 150
WebBrowser1.Height = Me.Height - 1630
cmbAddress.Width = Me.Width - 2800
Command5.Left = Me.Width - 1850
ProgressBar1.Left = Me.Width - 1250
Line1.X2 = Me.Width
Line2.X2 = Me.Width
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
    Select Case Button.Key
        Case "back"
            WebBrowser1.GoBack
        Case "forward"
            WebBrowser1.GoForward
        Case "refresh"
            WebBrowser1.Refresh
        Case "stop"
            On Error Resume Next
            WebBrowser1.Stop
    End Select
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
stats.Panels(1).Text = "Done"
End Sub

Private Sub WebBrowser1_DownloadComplete()
stats.Panels(1).Text = "Done"
End Sub

Private Sub WebBrowser1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
WebBrowser1.Navigate (Source)
End Sub

Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
Cancel = True
End Sub

Private Sub WebBrowser1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error Resume Next
ProgressBar1.Max = ProgressMax
ProgressBar1.Value = Progress
End Sub

Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
stats.Panels(1).Text = Text
End Sub

Private Sub WebBrowser1_TitleChange(ByVal Text As String)
Me.Caption = Text & " - Glycerine Browser"
End Sub
