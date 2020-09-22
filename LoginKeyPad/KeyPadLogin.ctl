VERSION 5.00
Begin VB.UserControl KeyPadLogin 
   ClientHeight    =   1335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2040
   HitBehavior     =   2  'Use Paint
   Picture         =   "KeyPadLogin.ctx":0000
   ScaleHeight     =   1335
   ScaleWidth      =   2040
   Begin VB.TextBox Text1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   105
      TabIndex        =   2
      Text            =   "6"
      Top             =   1875
      Width           =   1935
   End
   Begin VB.TextBox lblOutput 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1020
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   60
      Width           =   960
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Clear"
      Height          =   210
      Left            =   465
      TabIndex        =   5
      Top             =   990
      Width           =   405
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      Height          =   210
      Left            =   1215
      TabIndex        =   4
      Top             =   540
      Width           =   585
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter"
      Height          =   210
      Left            =   1290
      TabIndex        =   3
      Top             =   990
      Width           =   495
   End
   Begin VB.Image Image22 
      Height          =   315
      Left            =   1020
      Picture         =   "KeyPadLogin.ctx":18F06
      Top             =   495
      Width           =   945
   End
   Begin VB.Label Password 
      Caption         =   "23446"
      Height          =   30
      Left            =   7725
      TabIndex        =   1
      Top             =   6420
      Width           =   120
   End
   Begin VB.Image Image9 
      Height          =   315
      Left            =   645
      Picture         =   "KeyPadLogin.ctx":19F08
      Top             =   645
      Width           =   300
   End
   Begin VB.Image Image8 
      Height          =   315
      Left            =   375
      Picture         =   "KeyPadLogin.ctx":1A436
      Top             =   645
      Width           =   285
   End
   Begin VB.Image Image7 
      Height          =   315
      Left            =   90
      Picture         =   "KeyPadLogin.ctx":1A964
      Top             =   645
      Width           =   285
   End
   Begin VB.Image Image5 
      Height          =   315
      Left            =   360
      Picture         =   "KeyPadLogin.ctx":1AE92
      Top             =   345
      Width           =   300
   End
   Begin VB.Image Image6 
      Height          =   315
      Left            =   645
      Picture         =   "KeyPadLogin.ctx":1B3C0
      Top             =   345
      Width           =   300
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   645
      Picture         =   "KeyPadLogin.ctx":1B8EE
      Top             =   60
      Width           =   300
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   360
      Picture         =   "KeyPadLogin.ctx":1BDE0
      Top             =   60
      Width           =   300
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   90
      Picture         =   "KeyPadLogin.ctx":1C2D2
      Top             =   60
      Width           =   285
   End
   Begin VB.Image Image4 
      Height          =   315
      Left            =   90
      Picture         =   "KeyPadLogin.ctx":1C7C4
      Top             =   345
      Width           =   285
   End
   Begin VB.Image Image19 
      Height          =   315
      Left            =   1020
      Picture         =   "KeyPadLogin.ctx":1CCF2
      Top             =   945
      Width           =   945
   End
   Begin VB.Image Image12 
      Height          =   300
      Left            =   645
      Picture         =   "KeyPadLogin.ctx":1DCF4
      Top             =   60
      Width           =   300
   End
   Begin VB.Image Image15 
      Height          =   315
      Left            =   645
      Picture         =   "KeyPadLogin.ctx":1E1E6
      Top             =   345
      Width           =   300
   End
   Begin VB.Image Image18 
      Height          =   315
      Left            =   645
      Picture         =   "KeyPadLogin.ctx":1E714
      Top             =   645
      Width           =   300
   End
   Begin VB.Image Image17 
      Height          =   315
      Left            =   360
      Picture         =   "KeyPadLogin.ctx":1EC42
      Top             =   645
      Width           =   300
   End
   Begin VB.Image Image16 
      Height          =   315
      Left            =   75
      Picture         =   "KeyPadLogin.ctx":1F170
      Top             =   645
      Width           =   300
   End
   Begin VB.Image Image13 
      Height          =   315
      Left            =   75
      Picture         =   "KeyPadLogin.ctx":1F69E
      Top             =   345
      Width           =   300
   End
   Begin VB.Image Image14 
      Height          =   315
      Left            =   360
      Picture         =   "KeyPadLogin.ctx":1FBCC
      Top             =   345
      Width           =   300
   End
   Begin VB.Image Image11 
      Height          =   300
      Left            =   360
      Picture         =   "KeyPadLogin.ctx":200FA
      Top             =   60
      Width           =   300
   End
   Begin VB.Image Image10 
      Height          =   300
      Left            =   75
      Picture         =   "KeyPadLogin.ctx":205EC
      Top             =   60
      Width           =   300
   End
   Begin VB.Image Image21 
      Height          =   315
      Left            =   1020
      Picture         =   "KeyPadLogin.ctx":20ADE
      Top             =   945
      Width           =   930
   End
   Begin VB.Image Image20 
      Height          =   315
      Left            =   1020
      Picture         =   "KeyPadLogin.ctx":21A8C
      Top             =   495
      Width           =   930
   End
   Begin VB.Image Image24 
      Height          =   300
      Left            =   75
      Picture         =   "KeyPadLogin.ctx":22A3A
      Top             =   960
      Width           =   300
   End
   Begin VB.Image Image23 
      Height          =   285
      Left            =   75
      Picture         =   "KeyPadLogin.ctx":22F2C
      Top             =   960
      Width           =   300
   End
   Begin VB.Image Image25 
      Height          =   315
      Left            =   360
      Picture         =   "KeyPadLogin.ctx":233E2
      Top             =   945
      Width           =   585
   End
   Begin VB.Image Image26 
      Height          =   315
      Left            =   360
      Picture         =   "KeyPadLogin.ctx":23DFC
      Top             =   945
      Width           =   585
   End
End
Attribute VB_Name = "KeyPadLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Dim x1 As Long, y1 As Long

'Default Property Values:
Const m_def_Enabled = True
'Const m_def_Enabled = True
'Property Variables:
Dim m_Enabled As Boolean

Option Explicit
Public Event Denied()
Public Event Granted()
Public LoginSucceeded As Boolean

Private Sub Image19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image19.Visible = False
Image21.Visible = True
End Sub

Private Sub Image19_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image21.Visible = False
Image19.Visible = True
'check for correct password
    If lblOutput = Password.Caption Then
        'place code to here to pass the
        'success to the calling sub
        'setting a global var is the easiest
        LoginSucceeded = True
        RaiseEvent Granted
    Else
        MsgBox "Invalid Password, try again!", , "Login"
        SendKeys "{Home}+{End}"
        lblOutput.Text = ""
    End If
End Sub

Private Sub Image22_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image22.Visible = False
Image20.Visible = True
 

End Sub

Private Sub Image22_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image20.Visible = False
Image22.Visible = True
'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    RaiseEvent Denied
End Sub







Private Sub Image24_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image24.Visible = False
Image23.Visible = True
End Sub

Private Sub Image24_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image23.Visible = False
Image24.Visible = True
lblOutput.Text = lblOutput.Text & "0"
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image19.Visible = False
Image21.Visible = True
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image21.Visible = False
Image19.Visible = True
'check for correct password
    If lblOutput = Password.Caption Then
        'place code to here to pass the
        'success to the calling sub
        'setting a global var is the easiest
        LoginSucceeded = True
        RaiseEvent Granted
    Else
        MsgBox "Invalid Password, try again!", , "Login"
        SendKeys "{Home}+{End}"
        lblOutput.Text = ""
    End If
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image22.Visible = False
Image20.Visible = True
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image20.Visible = False
Image22.Visible = True
'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    RaiseEvent Denied
End Sub



Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image25.Visible = False
Image26.Visible = True
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image26.Visible = False
Image25.Visible = True
lblOutput.Text = ""
End Sub

Private Sub lblOutput_Change()

If Len(lblOutput.Text) > Text1.Text Then 'Limit of input is - 33 numbers max
    lblOutput.Text = Left$(lblOutput.Text, Text1.Text)
    End If
    

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Visible = False
Image10.Visible = True
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image10.Visible = False
Image1.Visible = True
lblOutput.Text = lblOutput.Text & "7"
End Sub



Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = False
Image11.Visible = True
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image11.Visible = False
Image2.Visible = True
lblOutput.Text = lblOutput.Text & "8"
End Sub



Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Visible = False
Image12.Visible = True

End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image12.Visible = False
Image3.Visible = True
lblOutput.Text = lblOutput.Text & "9"
End Sub


Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Visible = False
Image13.Visible = True

End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image13.Visible = False
Image4.Visible = True
lblOutput.Text = lblOutput.Text & "4"
End Sub


Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.Visible = False
Image14.Visible = True

End Sub

Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image14.Visible = False
Image5.Visible = True
lblOutput.Text = lblOutput.Text & "5"
End Sub

Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image6.Visible = False
Image15.Visible = True

End Sub

Private Sub Image6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image15.Visible = False
Image6.Visible = True
lblOutput.Text = lblOutput.Text & "6"
End Sub

Private Sub Image7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image7.Visible = False
Image16.Visible = True
End Sub

Private Sub Image7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image16.Visible = False
Image7.Visible = True
lblOutput.Text = lblOutput.Text & "1"
End Sub
Private Sub Image8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image8.Visible = False
Image17.Visible = True

End Sub

Private Sub Image8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image17.Visible = False
Image8.Visible = True
lblOutput.Text = lblOutput.Text & "2"
End Sub


Private Sub Image9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image9.Visible = False
Image18.Visible = True

End Sub

Private Sub Image9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image18.Visible = False
Image9.Visible = True
lblOutput.Text = lblOutput.Text & "3"
End Sub

Private Sub lblOutput_KeyPress(KeyAscii As Integer)
If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyInsert Then
KeyAscii = 0
End If

End Sub

Private Sub UserControl_Paint()
UserControl.Width = 2040
UserControl.Height = 1335
End Sub
Public Property Get Caption() As String
    Caption = Password.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Password.Caption() = New_Caption
    PropertyChanged "Password"
End Property
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Password", Password.Caption, "password")
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Password.Caption = PropBag.ReadProperty("Password", "password")
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
End Sub
