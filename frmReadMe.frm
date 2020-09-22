VERSION 5.00
Begin VB.Form frmReadMe 
   Caption         =   " READ ME"
   ClientHeight    =   9705
   ClientLeft      =   4500
   ClientTop       =   795
   ClientWidth     =   6390
   Icon            =   "frmReadMe.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9705
   ScaleWidth      =   6390
   Begin VB.CommandButton Command1 
      Caption         =   "&QUIT"
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   9120
      Width           =   1215
   End
   Begin VB.TextBox txtReadMe 
      Height          =   3975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   3720
      Width           =   6135
   End
   Begin VB.Label Label3 
      Caption         =   "Label1"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   8280
      Width           =   6255
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   7800
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   7740
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   3615
      Left            =   120
      Picture         =   "frmReadMe.frx":0CCA
      Stretch         =   -1  'True
      ToolTipText     =   "CAVER DAVE SELF PORTRAIT"
      Top             =   0
      Width           =   6135
   End
End
Attribute VB_Name = "frmReadMe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
'**************************
'   and centre the form   *
'**************************
Dim TopCorner As Integer
  Dim LeftCorner As Integer
  'centres the form on the screen
  If Me.WindowState <> 0 Then Exit Sub

  TopCorner = (Screen.Height - Me.Height) \ 2
  LeftCorner = (Screen.Width - Me.Width) \ 2
  Me.Move LeftCorner, TopCorner
  
  Label1.Caption = App.ProductName
  Label2.Caption = "Version: " & App.Major & ". " & App.Minor & ". " & App.Revision
  Label3.Caption = App.Comments
  
  Call bkReadMe
  
End Sub
Private Sub bkReadMe()
Dim Filehandle As Integer
  Dim FileLength
Dim var1
  Filehandle = FreeFile

Open App.Path & "\ReadMe.txt" For Input As #Filehandle
FileLength = LOF(Filehandle)
var1 = Input(FileLength, #Filehandle)
txtReadMe.Text = var1
Close #Filehandle
End Sub

