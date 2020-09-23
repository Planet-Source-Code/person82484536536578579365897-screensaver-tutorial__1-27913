VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Blanker"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1680
      Top             =   2040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   10
      X2              =   3240
      Y1              =   1200
      Y2              =   1200
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bWhite As Boolean
Private Sub Form_Activate()
Line1.X1 = frmMain.Width \ 2
Line1.Y1 = frmMain.Height \ 2
Line1.X2 = frmMain.Width \ 2
Line1.Y2 = frmMain.Height \ 2
Timer1.Enabled = True
End Sub

Private Sub Form_Click()
If GetSetting("Blanker", "Settings", "Password", "") <> "" Then
frmPassword.Show 'If a password is set then show the password box
Else
exitScreensaver 'exit the screensaver
End If
End Sub

Private Sub Form_DblClick()
If GetSetting("Blanker", "Settings", "Password", "") <> "" Then
frmPassword.Show 'If a password is set then show the password box
Else
exitScreensaver 'exit the screensaver
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If GetSetting("Blanker", "Settings", "Password", "") <> "" Then
frmPassword.Show 'If a password is set then show the password box
Else
exitScreensaver 'exit the screensaver
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If GetSetting("Blanker", "Settings", "Password", "") <> "" Then
frmPassword.Show 'If a password is set then show the password box
Else
exitScreensaver 'exit the screensaver
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If GetSetting("Blanker", "Settings", "Password", "") <> "" Then
frmPassword.Show 'If a password is set then show the password box
Else
exitScreensaver 'exit the screensaver
End If
End Sub

Private Sub Form_Load()
ShowCursor False 'Hide the cursor
bWhite = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If GetSetting("Blanker", "Settings", "Password", "") <> "" Then
frmPassword.Show 'If a password is set then show the password box
Else
exitScreensaver 'exit the screensaver
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If GetSetting("Blanker", "Settings", "Password", "") <> "" Then
frmPassword.Show 'If a password is set then show the password box
Else
exitScreensaver 'exit the screensaver
End If
End Sub

Private Sub Timer1_Timer()
Line1.BorderWidth = Line1.BorderWidth + 3 'Increase the border width by 3
If Line1.BorderWidth >= 1000 Then
Line1.BorderWidth = 1 'If the line's border width _
gets bigger than the screen then set it back to one.
If bWhite = True Then
Line1.BorderColor = vbGreen
bWhite = False
GoTo This
End If
If bWhite = False Then
Line1.BorderColor = vbWhite
bWhite = True
End If
End If
This:
End Sub
