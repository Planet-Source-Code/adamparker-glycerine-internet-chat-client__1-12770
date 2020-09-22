VERSION 5.00
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "Glycerine - The Ultimate Free Internet Chat"
   ClientHeight    =   6810
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9615
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   60
      Left            =   0
      ScaleHeight     =   60
      ScaleWidth      =   9615
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   2640
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   2640
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSetAction 
         Caption         =   "&Set Custom Action"
      End
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuFileView 
      Caption         =   "View"
      Begin VB.Menu mnuFileDebug 
         Caption         =   "&Debug Window"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuFileTime 
         Caption         =   "&Time Stamp"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuViewLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConnect 
         Caption         =   "&Connect Screen"
         Shortcut        =   ^N
      End
   End
   Begin VB.Menu mnuSPGC 
      Caption         =   "My Glycerine"
      Begin VB.Menu mnuExWeb 
         Caption         =   "&Glycerine Web Browser"
      End
      Begin VB.Menu mnuEditSvr 
         Caption         =   "Edit &Server List"
      End
      Begin VB.Menu mnuFilePrefs 
         Caption         =   "&Preferences"
      End
   End
   Begin VB.Menu mnuProg 
      Caption         =   "&Programs"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&Edit Pad"
      End
   End
   Begin VB.Menu mnu_Window 
      Caption         =   "&Window"
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTile 
         Caption         =   "&Tile Horizontally"
      End
      Begin VB.Menu mnuWindowArrange 
         Caption         =   "&Tile Vertically"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuChatc 
         Caption         =   "&Chat commands"
      End
      Begin VB.Menu mnuHelpLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmQuit.Show
End Sub

Private Sub MDIForm_Resize()
Line1.X2 = Me.Width
Line2.X2 = Me.Width
End Sub

Private Sub mnuChatc_Click()
frmCommands.Show
End Sub

Private Sub mnuConnect_Click()
frmConnect.Show
End Sub

Private Sub mnuEditSvr_Click()
frmServerList.Show
End Sub

Private Sub mnuExWeb_Click()
Dim Popup As New frmPopUp
Popup.Show
End Sub

Private Sub mnuFileDebug_Click()
frmDebug.Show
End Sub

Private Sub mnuFileExit_Click()
Unload Me
End Sub

Private Sub mnuFileNew_Click()
frmEdit.Show
End Sub

Private Sub mnuFilePrefs_Click()
frmPref.Show
End Sub

Private Sub mnuFileTime_Click()
If mnuFileTime.Checked = True Then
mnuFileTime.Checked = False
Else
mnuFileTime.Checked = True
End If
End Sub

Private Sub mnuHelpAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuScript_Click()
frmScript.Show
End Sub

Private Sub mnuSetAction_Click()
frmCmd.Show
End Sub

Private Sub mnuStock_Click()
Form1.Show
End Sub

Private Sub mnuWindowArrange_Click()
    mdiMain.Arrange vbTileVertical
End Sub

Private Sub mnuWindowCascade_Click()
    mdiMain.Arrange vbCascade
End Sub

Private Sub mnuWindowTile_Click()
    mdiMain.Arrange vbTileHorizontal
End Sub

