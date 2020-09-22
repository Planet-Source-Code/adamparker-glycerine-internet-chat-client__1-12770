VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "frmMenu"
   ClientHeight    =   255
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   3840
   LinkTopic       =   "Form2"
   ScaleHeight     =   255
   ScaleWidth      =   3840
   Visible         =   0   'False
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy Text"
      End
      Begin VB.Menu mnuEditLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save Chat Text"
      End
      Begin VB.Menu mnuEditClear 
         Caption         =   "C&lear Chat Text"
      End
   End
   Begin VB.Menu mnuOpt 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptSlap 
         Caption         =   "&Slap"
      End
      Begin VB.Menu mnuLOL 
         Caption         =   "&LOL at "
      End
      Begin VB.Menu mnuScare 
         Caption         =   "&Scare "
      End
      Begin VB.Menu mnuPunch 
         Caption         =   "&Punch"
      End
      Begin VB.Menu mnuSay 
         Caption         =   "&Custom Action"
      End
      Begin VB.Menu mnuOptline1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptPM 
         Caption         =   "&Send PM to "
      End
      Begin VB.Menu mnuCust 
         Caption         =   "&Set Custom Action"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuCust_Click()
frmCmd.Show
End Sub

Private Sub mnuEditClear_Click()
If MsgBox("Are you sure you want to clear that chat room text?", vbYesNo) = vbYes Then frmMain.txtText = ""
End Sub

Private Sub mnuEditCopy_Click()
Call Clipboard.SetText(frmMain.txtText.SelText, 1)
End Sub
Private Sub mnuFileSave_Click()
Call frmMain.dlgColor.ShowSave
Call frmMain.txtText.SaveFile(frmMain.dlgColor.FileName, 1)
End Sub

Private Sub mnuLOL_Click()
frmConnect.winsck.SendData ("me-* " & frmConnect.txtName & " is laughing out loud at " & frmMain.lstUsers.Text)
End Sub

Private Sub mnuOptPM_Click()
frmMSG.Show
frmMSG.txtWho = frmMain.lstUsers.Text
End Sub

Private Sub mnuOptSlap_Click()
frmConnect.winsck.SendData ("me-* " & frmConnect.txtName & " slaps " & frmMain.lstUsers.Text & " around a bit with a large trout")
End Sub

Private Sub mnuPunch_Click()
frmConnect.winsck.SendData ("me-* " & frmConnect.txtName & " punches " & frmMain.lstUsers.Text & " and knocks him out")
End Sub

Private Sub mnuSay_Click()
On Error GoTo err:
strData = ReadFile(App.Path & "\prefs\cust.dat")
strData = Replace(strData, "%n", frmMain.lstUsers.Text)
strData = Replace(strData, "%t", Time)
strData = Replace(strData, "%d", Date)
strData = Replace(strData, "%v", App.Major & "." & App.Minor & "." & App.Revision)
strData = Replace(strData, "%r", Random(1, 1000))
strData = Replace(strData, "%s", frmConnect.cmbServer)
strData = Replace(strData, "%ip", frmConnect.winsck.LocalIP)
strData = Replace(strData, "%p", frmConnect.txtName)
frmConnect.winsck.SendData ("me- *" & frmConnect.txtName & ": " & strData)
Exit Sub
err:
End Sub

Private Sub mnuScare_Click()
frmConnect.winsck.SendData ("me-* " & frmConnect.txtName & " sneaks up behind " & frmMain.lstUsers.Text & " and gives him a fright")
End Sub
