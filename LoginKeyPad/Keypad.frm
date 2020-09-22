VERSION 5.00
Begin VB.Form Keypad 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1320
   ClientLeft      =   3495
   ClientTop       =   1785
   ClientWidth     =   2055
   ControlBox      =   0   'False
   Icon            =   "Keypad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   2055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Loginkeypad.KeyPadLogin KeyPadLogin1 
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   2355
      Password        =   "23446"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "Keypad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub KeyPadLogin1_Denied()
End
End Sub

Private Sub KeyPadLogin1_Granted()
End
End Sub

