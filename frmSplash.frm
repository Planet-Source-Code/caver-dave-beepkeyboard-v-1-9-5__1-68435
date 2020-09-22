VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4335
   ClientLeft      =   4980
   ClientTop       =   3420
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   600
      Top             =   2160
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   20
      Height          =   4095
      Left            =   120
      Top             =   120
      Width           =   5775
   End
   Begin VB.Image Image2 
      Height          =   2250
      Left            =   375
      Picture         =   "frmSplash.frx":0000
      Top             =   480
      Width           =   5250
   End
   Begin VB.Image Image1 
      Height          =   945
      Left            =   345
      Picture         =   "frmSplash.frx":0909
      Top             =   2910
      Width           =   5250
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   345
      TabIndex        =   0
      Top             =   3120
      Width           =   5250
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Private Sub Form_Load()
'**************************
'      play sig tune      *
'   and centre the form   *
'**************************
Dim TopCorner As Integer
  Dim LeftCorner As Integer
  'centres the form on the screen
  If Me.WindowState <> 0 Then Exit Sub

  TopCorner = (Screen.Height - Me.Height) \ 2
  LeftCorner = (Screen.Width - Me.Width) \ 2
  Me.Move LeftCorner, TopCorner
Call bkBeep
Timer1.Enabled = True
Timer1.Interval = 2000
End Sub
Private Sub bkBeep()
'**************************
'         sig tune        *
'**************************
Beep 783.99, 138
Beep 783.99, 124
Beep 783.99, 125
Beep 622.25, 1983
Beep 0, 123
Beep 698.46, 137
Beep 698.46, 124
Beep 698.46, 127
Beep 587.33, 2493

End Sub

Private Sub Timer1_Timer()
'********************************
'  splash screen display timer  *
'********************************
If Timer1.Interval = 2000 Then
Unload Me
frmMain.Show
Timer1.Enabled = False
End If
End Sub
