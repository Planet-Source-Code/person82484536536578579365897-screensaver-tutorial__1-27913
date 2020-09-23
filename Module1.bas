Attribute VB_Name = "Module1"
Option Explicit 'All variables must be declared
'Constants
Public Const SW_SHOWNORMAL = 1
Private Const APP_NAME = "Blanker"
'API's
'This Function Shows and hides the cursor.
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
'This function finds if another instance of the saver is running.
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVallpClassName As String, ByVal lpWindowName As String) As Long

Sub Main()
'This sub is called when windows wants the screensaver to run a certain event.
'This locates what windows wants and activates it.
Select Case Mid(UCase$(Trim$(Command$)), 1, 2)

Case "/C" 'Configurations mode called
frmSetup.Show 1

Case "", "/S" 'Screensaver mode
runScreensaver

Case "/A" 'Password protect dialog
frmPassSetup.Show 1
Case "/P" 'Preview mode
'The preview mode is very advanced. It is when you see a clip of it on the little 'monitor. Just leave the monitor screen blank.
End
End Select
End Sub

Private Sub runScreensaver() 'Run the screen saver
checkInstance 'Make sure no other instances are running
ShowCursor False 'Disable cursor
'load Screen Saver's main form
Load frmMain
frmMain.Show
End Sub


Private Sub checkInstance()
'If no previous instance is running, exit sub
If Not App.PrevInstance Then Exit Sub

'check for another instance of screen saver
If FindWindow(vbNullString, APP_NAME) Then Exit Sub

'Set our caption so other instances can find
'us in the previous line.

frmMain.Caption = APP_NAME
End Sub

Sub exitScreensaver() 'Exit the screensaver
ShowCursor True 'Show the cursor
End
End Sub

