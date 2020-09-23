VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   465
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4185
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   465
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Created By: Aaron Lindsay"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4140
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
