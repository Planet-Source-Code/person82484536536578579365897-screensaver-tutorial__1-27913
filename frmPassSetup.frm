VERSION 5.00
Begin VB.Form frmPassSetup 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Password"
   ClientHeight    =   915
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3630
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   915
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   960
      MaxLength       =   40
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmPassSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
SaveSetting "Blanker", "Settings", "Password", txtPassword.Text
Unload Me
End Sub

