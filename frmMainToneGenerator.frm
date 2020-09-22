VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   " BEEP KEYBOARD AND PLAYER"
   ClientHeight    =   5595
   ClientLeft      =   1440
   ClientTop       =   2070
   ClientWidth     =   12900
   Icon            =   "frmMainToneGenerator.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   12900
   Begin VB.TextBox Text1 
      Height          =   1335
      Left            =   3360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   6840
      TabIndex        =   115
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&OPEN TUNE"
      Height          =   375
      Left            =   11280
      TabIndex        =   114
      Top             =   3000
      Width           =   1455
   End
   Begin VB.ComboBox cmbTunelister 
      Height          =   315
      Left            =   9120
      MouseIcon       =   "frmMainToneGenerator.frx":0CCA
      MousePointer    =   99  'Custom
      TabIndex        =   113
      Text            =   "Tune Lister"
      Top             =   3030
      Width           =   2055
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&DELETE SELECTED NOTE"
      Height          =   855
      Left            =   9120
      TabIndex        =   112
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   9120
      TabIndex        =   111
      Text            =   "0"
      Top             =   3840
      Width           =   615
   End
   Begin VB.ListBox List3 
      Height          =   1815
      Left            =   7920
      TabIndex        =   25
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "READ &ME"
      Height          =   615
      Left            =   12000
      TabIndex        =   24
      Top             =   3960
      Width           =   735
   End
   Begin VB.Timer rptTim 
      Enabled         =   0   'False
      Left            =   3720
      Top             =   4080
   End
   Begin VB.Timer tmrRepeat 
      Enabled         =   0   'False
      Left            =   4320
      Top             =   4080
   End
   Begin VB.TextBox txtTuneSave 
      Height          =   285
      Left            =   9120
      TabIndex        =   17
      Text            =   "Tune Name"
      Top             =   3525
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&SAVE TUNE"
      Height          =   375
      Left            =   11280
      TabIndex        =   16
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "RESET &ALL"
      Height          =   615
      Left            =   11120
      TabIndex        =   15
      Top             =   4440
      Width           =   735
   End
   Begin VB.ListBox List2 
      Height          =   1815
      Left            =   6720
      TabIndex        =   12
      Top             =   3240
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   5520
      TabIndex        =   11
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtPlayLister 
      Height          =   285
      Left            =   4680
      TabIndex        =   9
      Text            =   "0"
      Top             =   4665
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&PLAY LIST"
      Height          =   615
      Left            =   10240
      TabIndex        =   8
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E&XIT"
      Height          =   375
      Left            =   12000
      TabIndex        =   3
      Top             =   4680
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "THE ALMOST(83 KEYS) FULL KEYBOARD"
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   12615
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   49
         Left            =   12060
         ScaleHeight     =   1485
         ScaleWidth      =   345
         TabIndex        =   109
         ToolTipText     =   "PAUSE"
         Top             =   840
         Width           =   375
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   0
         Left            =   11340
         ScaleHeight     =   1065
         ScaleWidth      =   105
         TabIndex        =   59
         ToolTipText     =   "3729.3"
         Top             =   840
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   1
         Left            =   11100
         ScaleHeight     =   1065
         ScaleWidth      =   105
         TabIndex        =   58
         ToolTipText     =   "3322.4"
         Top             =   840
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   2
         Left            =   10860
         ScaleHeight     =   1065
         ScaleWidth      =   105
         TabIndex        =   57
         ToolTipText     =   "2960"
         Top             =   840
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   3
         Left            =   10380
         ScaleHeight     =   1065
         ScaleWidth      =   105
         TabIndex        =   56
         ToolTipText     =   "2498"
         Top             =   840
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   4
         Left            =   10140
         ScaleHeight     =   1065
         ScaleWidth      =   105
         TabIndex        =   55
         ToolTipText     =   "2217.5"
         Top             =   840
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   5
         Left            =   9660
         ScaleHeight     =   1065
         ScaleWidth      =   105
         TabIndex        =   54
         ToolTipText     =   "1864.7"
         Top             =   840
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   6
         Left            =   9420
         ScaleHeight     =   1065
         ScaleWidth      =   105
         TabIndex        =   53
         ToolTipText     =   "1661.2"
         Top             =   840
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   7
         Left            =   9180
         ScaleHeight     =   1065
         ScaleWidth      =   105
         TabIndex        =   52
         ToolTipText     =   "1480"
         Top             =   840
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   8
         Left            =   8700
         ScaleHeight     =   1065
         ScaleWidth      =   105
         TabIndex        =   51
         ToolTipText     =   "1244.5"
         Top             =   840
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   9
         Left            =   8460
         ScaleHeight     =   1065
         ScaleWidth      =   105
         TabIndex        =   50
         ToolTipText     =   "1108.7"
         Top             =   840
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   10
         Left            =   7980
         ScaleHeight     =   1065
         ScaleWidth      =   105
         TabIndex        =   49
         ToolTipText     =   "932.33"
         Top             =   840
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   11
         Left            =   7740
         ScaleHeight     =   1065
         ScaleWidth      =   105
         TabIndex        =   48
         ToolTipText     =   "830.61"
         Top             =   840
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   12
         Left            =   7500
         ScaleHeight     =   1065
         ScaleWidth      =   105
         TabIndex        =   47
         ToolTipText     =   "739.99"
         Top             =   840
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   13
         Left            =   7020
         ScaleHeight     =   1065
         ScaleWidth      =   105
         TabIndex        =   46
         ToolTipText     =   "622.25"
         Top             =   840
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   14
         Left            =   6780
         ScaleHeight     =   1065
         ScaleWidth      =   105
         TabIndex        =   45
         ToolTipText     =   "554.37"
         Top             =   840
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   15
         Left            =   6300
         ScaleHeight     =   1065
         ScaleWidth      =   105
         TabIndex        =   44
         ToolTipText     =   "466.16"
         Top             =   840
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   16
         Left            =   6060
         ScaleHeight     =   1065
         ScaleWidth      =   105
         TabIndex        =   43
         ToolTipText     =   "415.3"
         Top             =   840
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   17
         Left            =   5820
         ScaleHeight     =   1065
         ScaleWidth      =   105
         TabIndex        =   42
         ToolTipText     =   "369.99"
         Top             =   840
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   18
         Left            =   5340
         ScaleHeight     =   1065
         ScaleWidth      =   105
         TabIndex        =   41
         ToolTipText     =   "311.13"
         Top             =   840
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   19
         Left            =   5100
         ScaleHeight     =   1065
         ScaleWidth      =   105
         TabIndex        =   40
         ToolTipText     =   "277.18"
         Top             =   840
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   20
         Left            =   4620
         ScaleHeight     =   1065
         ScaleWidth      =   105
         TabIndex        =   39
         ToolTipText     =   "233.08"
         Top             =   840
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   21
         Left            =   4380
         ScaleHeight     =   1065
         ScaleWidth      =   105
         TabIndex        =   38
         ToolTipText     =   "207.65"
         Top             =   840
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   22
         Left            =   4140
         ScaleHeight     =   1065
         ScaleWidth      =   105
         TabIndex        =   37
         ToolTipText     =   "185"
         Top             =   840
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   23
         Left            =   3660
         ScaleHeight     =   1065
         ScaleWidth      =   105
         TabIndex        =   36
         ToolTipText     =   "155.56"
         Top             =   840
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   24
         Left            =   3420
         ScaleHeight     =   1065
         ScaleWidth      =   105
         TabIndex        =   35
         ToolTipText     =   "138.59"
         Top             =   840
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   25
         Left            =   2940
         ScaleHeight     =   1065
         ScaleWidth      =   105
         TabIndex        =   34
         ToolTipText     =   "116.54"
         Top             =   840
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   26
         Left            =   2700
         ScaleHeight     =   1065
         ScaleWidth      =   105
         TabIndex        =   33
         ToolTipText     =   "103.83"
         Top             =   840
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   27
         Left            =   2460
         ScaleHeight     =   1065
         ScaleWidth      =   105
         TabIndex        =   32
         ToolTipText     =   "92.499"
         Top             =   840
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   28
         Left            =   1980
         ScaleHeight     =   1065
         ScaleWidth      =   105
         TabIndex        =   31
         ToolTipText     =   "77.782"
         Top             =   840
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   29
         Left            =   1740
         ScaleHeight     =   1065
         ScaleWidth      =   105
         TabIndex        =   30
         ToolTipText     =   "69.296"
         Top             =   840
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   30
         Left            =   1500
         ScaleHeight     =   1065
         ScaleWidth      =   105
         TabIndex        =   29
         ToolTipText     =   "58.27"
         Top             =   840
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   31
         Left            =   1020
         ScaleHeight     =   1065
         ScaleWidth      =   105
         TabIndex        =   28
         ToolTipText     =   "51.913"
         Top             =   840
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   32
         Left            =   780
         ScaleHeight     =   1065
         ScaleWidth      =   105
         TabIndex        =   27
         ToolTipText     =   "46.249"
         Top             =   840
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   33
         Left            =   300
         ScaleHeight     =   1065
         ScaleWidth      =   105
         TabIndex        =   26
         ToolTipText     =   "38.891"
         Top             =   840
         Width           =   135
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   46
         Left            =   600
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   60
         ToolTipText     =   "43.654"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   45
         Left            =   840
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   61
         ToolTipText     =   "48.999"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   44
         Left            =   1080
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   62
         ToolTipText     =   "55"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   43
         Left            =   1320
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   63
         ToolTipText     =   "61.735"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   39
         Left            =   2280
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   64
         ToolTipText     =   "87.307"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   38
         Left            =   2520
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   65
         ToolTipText     =   "97.999"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   37
         Left            =   2760
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   66
         ToolTipText     =   "110"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   36
         Left            =   3000
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   67
         ToolTipText     =   "123.47"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   32
         Left            =   3960
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   68
         ToolTipText     =   "174.61"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   31
         Left            =   4200
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   69
         ToolTipText     =   "196"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   30
         Left            =   4440
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   70
         ToolTipText     =   "220"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   29
         Left            =   4680
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   71
         ToolTipText     =   "246.94"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   33
         Left            =   3720
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   72
         ToolTipText     =   "164.8"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   34
         Left            =   3480
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   73
         ToolTipText     =   "146.83"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   35
         Left            =   3240
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   74
         ToolTipText     =   "130.81"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   40
         Left            =   2040
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   75
         ToolTipText     =   "82.407"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   41
         Left            =   1800
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   76
         ToolTipText     =   "73.416"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   42
         Left            =   1560
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   77
         ToolTipText     =   "65.406"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   47
         Left            =   360
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   78
         ToolTipText     =   "41.203"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   48
         Left            =   120
         Picture         =   "frmMainToneGenerator.frx":0E1C
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   79
         ToolTipText     =   "36.708"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   25
         Left            =   5640
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   80
         ToolTipText     =   "349.23"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   24
         Left            =   5880
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   81
         ToolTipText     =   "392"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   23
         Left            =   6120
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   82
         ToolTipText     =   "440"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   22
         Left            =   6360
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   83
         ToolTipText     =   "493.38"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   18
         Left            =   7320
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   84
         ToolTipText     =   "698.46"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   17
         Left            =   7560
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   85
         ToolTipText     =   "783.99"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   16
         Left            =   7800
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   86
         ToolTipText     =   "880"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   15
         Left            =   8040
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   87
         ToolTipText     =   "987.77"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   11
         Left            =   9000
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   88
         ToolTipText     =   "1396.9"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   10
         Left            =   9240
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   89
         ToolTipText     =   "1568"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   12
         Left            =   8760
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   90
         ToolTipText     =   "1318.5"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   13
         Left            =   8520
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   91
         ToolTipText     =   "1174.7"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   14
         Left            =   8280
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   92
         ToolTipText     =   "1046.5"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   19
         Left            =   7080
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   93
         ToolTipText     =   "659.26"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   20
         Left            =   6840
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   94
         ToolTipText     =   "587.33"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   21
         Left            =   6600
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   95
         ToolTipText     =   "523.25"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   26
         Left            =   5400
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   96
         ToolTipText     =   "329.63"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   27
         Left            =   5160
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   97
         ToolTipText     =   "293.67"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   28
         Left            =   4920
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   98
         ToolTipText     =   "261.6"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   9
         Left            =   9480
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   99
         ToolTipText     =   "1760"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   8
         Left            =   9720
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   100
         ToolTipText     =   "1975.5"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   7
         Left            =   9960
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   101
         ToolTipText     =   "2093"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   6
         Left            =   10200
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   102
         ToolTipText     =   "2349.3"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   5
         Left            =   10440
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   103
         ToolTipText     =   "2637"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   4
         Left            =   10680
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   104
         ToolTipText     =   "2793.8"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   3
         Left            =   10920
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   105
         ToolTipText     =   "3136"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   2
         Left            =   11160
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   106
         ToolTipText     =   "3520"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   1
         Left            =   11400
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   107
         ToolTipText     =   "3951.1"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   0
         Left            =   11640
         ScaleHeight     =   1485
         ScaleWidth      =   225
         TabIndex        =   108
         ToolTipText     =   "4186"
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "THE NOT AT ALL A YAMAHA!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3773
         TabIndex        =   2
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "TONE DURATION"
      Height          =   1455
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   3135
      Begin VB.HScrollBar hsDur 
         Height          =   255
         LargeChange     =   100
         Left            =   120
         Max             =   10000
         Min             =   5
         SmallChange     =   5
         TabIndex        =   7
         Top             =   240
         Value           =   5
         Width           =   2895
      End
      Begin VB.TextBox txtDuration 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&RESET TONE DURATION"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   2415
      End
   End
   Begin VB.TextBox txtRepeatTime 
      Height          =   285
      Left            =   2040
      TabIndex        =   18
      Text            =   "0"
      Top             =   4545
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "NO REPEAT"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "NOTES"
      Height          =   495
      Index           =   4
      Left            =   7920
      TabIndex        =   110
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label LblLink 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "sds-software-maker.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   360
      Left            =   120
      TabIndex        =   20
      ToolTipText     =   "IF YOU LIKE THIS PROGRAM VISIT MY WEB SITE"
      Top             =   5040
      Width           =   2670
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   495
      Left            =   1800
      TabIndex        =   23
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "NOTE COUNT"
      Height          =   255
      Left            =   3480
      TabIndex        =   22
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "DURATION"
      Height          =   495
      Index           =   2
      Left            =   6720
      TabIndex        =   14
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "FREQUENCY"
      Height          =   495
      Index           =   1
      Left            =   5520
      TabIndex        =   13
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "MIDI NOTATION LIST"
      Height          =   495
      Index           =   0
      Left            =   3360
      TabIndex        =   10
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "TUNE TIME"
      Height          =   255
      Index           =   3
      Left            =   720
      TabIndex        =   21
      Top             =   4560
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const LB_GETTOPINDEX = &H18E
Private Const LB_SETTOPINDEX = &H197

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Const myerrfilepath = 53
Const mydeleteerr = 5
'Private Sub Check1_Click()
''**************************
''       repeat tune       *
''**************************
'If Check1.Value = 0 Then
'Check1.Caption = "NO REPEAT"
'ElseIf Check1.Value = 1 Then
'Check1.Caption = "REPEAT"
'tmrRepeat.Enabled = True
'tmrRepeat.Interval = Val(txtRepeatTime.Text)
'End If
'End Sub

Private Sub Command1_Click()
'**************************
'         end app         *
'**************************
Unload Me
Unload frmReadMe
End
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'**************************
'    set control cursor   *
'**************************
Command1.MousePointer = 99
Command1.MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub Command2_Click()
'**************************
' reset the tone duration *
'**************************
hsDur.Value = 250

End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'**************************
'    set control cursor   *
'**************************
Command2.MousePointer = 99
Command2.MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub Command3_Click()
'**************************
'      play the tune      *
'**************************
Dim Index As Integer
For Index = 0 To txtPlayLister.Text - 1
Beep List1.List(Index), List2.List(Index)
List1.Selected(Index) = True
List2.Selected(Index) = True
List3.Selected(Index) = True
Next Index
'If Check1.Value = 0 Then
'tmrRepeat.Enabled = False
'ElseIf Check1.Value = 1 Then
'Call Check1_Click
'End If
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'**************************
'    set control cursor   *
'**************************
Command3.MousePointer = 99
Command3.MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub Command4_Click()
'**************************
'         reset all       *
'**************************
hsDur.Value = 250
List1.Clear
List2.Clear
List3.Clear
txtPlayLister.Text = "0"
Text1.Text = ""
txtTuneSave.Text = "Tune Name"
cmbTunelister.Text = "Tune Lister"
Text2.Text = "0"
txtRepeatTime.Text = "0"
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'**************************
'    set control cursor   *
'**************************
Command4.MousePointer = 99
Command4.MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub Command5_Click()
'*********************************
'      save tune routines        *
'  *.bkt / *.bkf / *.bkd / *.bkn *
'*********************************
Call bkSaveFreq  '  *.bkf
Call bkSaveDur   '  *.bkd
Call bkSaveList  '  *.bkt
Call bkSaveNote '  *.bkn
End Sub

Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'**************************
'    set control cursor   *
'**************************
Command5.MousePointer = 99
Command5.MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub Command6_Click()
'******************************
'      open tune routines     *
'    *.bkf / *.bkd / *.bkn    *
'******************************
On Error GoTo fubar
  'opens the selected file for editing'
  Dim Msg As String
List1.Clear
List2.Clear
List3.Clear
Call bkOpenFreq '  *.bkf
Call bkOpenDur  '  *.bkd
Call bkOpenNote '  *.bkn
txtPlayLister.Text = List1.ListCount

fubar:
  If (Err.Number = myerrfilepath) Then
    Msg = UCase("you must select a file to open")
    If MsgBox(Msg) = vbOK Then
     frmMain.SetFocus
    End If
  End If
  Exit Sub
End Sub

Private Sub Command6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'**************************
'    set control cursor   *
'**************************
Command6.MousePointer = 99
Command6.MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub Command7_Click()
'**************************
'  display read me form   *
'**************************
frmReadMe.Show
End Sub

Private Sub Command8_Click()
'**************************
' tune editing as you go  *
'**************************
On Error GoTo fubar
  'nothing selected for editing'
  Dim Msg As String
List1.RemoveItem Text2.Text
List2.RemoveItem Text2.Text
List3.RemoveItem Text2.Text
txtPlayLister.Text = List1.ListCount

'Text3.Text = List2.List(Text2.Text)

txtRepeatTime.Text = Val(txtRepeatTime.Text) - Val(Text3.Text) ' subtract the deleted time from the tune time
fubar:
  If (Err.Number = mydeleteerr) Then
    Msg = UCase("you must have begun a tune")
    If MsgBox(Msg) = vbOK Then
     frmMain.SetFocus
    End If
  End If
  Exit Sub
End Sub

Private Sub Form_Load()
'**************************
'  set the tone duration  *
'   and centre the form   *
'**************************
Dim TopCorner As Integer
  Dim LeftCorner As Integer
  'centres the form on the screen
  If Me.WindowState <> 0 Then Exit Sub

  TopCorner = (Screen.Height - Me.Height) \ 2
  LeftCorner = (Screen.Width - Me.Width) \ 2
  Me.Move LeftCorner, TopCorner
  
Call bkOpenList ' open the tune lister file

hsDur.Value = 250
txtDuration.Text = hsDur.Value
End Sub

Private Sub hsDur_Change()
'***************************
' change the tone duration *
'***************************
txtDuration.Text = hsDur.Value
End Sub

Private Sub hsDur_Scroll()
'**************************
'    set control cursor   *
'**************************
hsDur.MousePointer = 99
hsDur.MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub List1_Click()
'*******************************************
'      synchs the listbox selections       *
'*******************************************
Text2.Text = List1.ListIndex ' gets the index of the selected item
Text3.Text = List2.List(Text2.Text) ' gets the duration of the selected item
List1.Selected(Text2.Text) = True
List2.Selected(Text2.Text) = True
List3.Selected(Text2.Text) = True

End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'**************************
'    set control cursor   *
'**************************
List1.MousePointer = 99
List1.MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub List2_Click()
'*******************************************
'      synchs the listbox selections       *
'*******************************************
Text2.Text = List2.ListIndex ' gets the index of the selected item

List1.Selected(Text2.Text) = True
List2.Selected(Text2.Text) = True
List3.Selected(Text2.Text) = True
End Sub

Private Sub List2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'**************************
'    set control cursor   *
'**************************
List2.MousePointer = 99
List2.MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub List3_Click()
'*******************************************
'      synchs the listbox selections       *
'*******************************************
Text2.Text = List3.ListIndex ' gets the index of the selected item

List1.Selected(Text2.Text) = True
List2.Selected(Text2.Text) = True
List3.Selected(Text2.Text) = True
End Sub

Private Sub List3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'**************************
'    set control cursor   *
'**************************
List3.MousePointer = 99
List3.MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub Picture1_Click(Index As Integer)
'*******************************************
'        play the treble keys (white)      *
'   adds the note info to the listboxes    *
'     totals the number of notes and       *
'       totals the compostion time         *
'*******************************************
Select Case Index
Case 0
Beep 4186#, txtDuration.Text
Text1.Text = Text1.Text & "108 "
List1.AddItem 4186#
List2.AddItem txtDuration.Text
List3.AddItem "C8"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 1
Beep 3951.1, txtDuration.Text
Text1.Text = Text1.Text & "107 "
List1.AddItem 3951.1
List2.AddItem txtDuration.Text
List3.AddItem "B7"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 2
Beep 3520#, txtDuration.Text
Text1.Text = Text1.Text & "105 "
List1.AddItem 3520#
List2.AddItem txtDuration.Text
List3.AddItem "A7"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 3
Beep 3136#, txtDuration.Text
Text1.Text = Text1.Text & "103 "
List1.AddItem 3136#
List2.AddItem txtDuration.Text
List3.AddItem "G7"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 4
Beep 2793.8, txtDuration.Text
Text1.Text = Text1.Text & "101 "
List1.AddItem 2793.8
List2.AddItem txtDuration.Text
List3.AddItem "F7"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 5
Beep 2637#, txtDuration.Text
Text1.Text = Text1.Text & "100 "
List1.AddItem 2637#
List2.AddItem txtDuration.Text
List3.AddItem "E7"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 6
Beep 2349.3, txtDuration.Text
Text1.Text = Text1.Text & "98 "
List1.AddItem 2349.3
List2.AddItem txtDuration.Text
List3.AddItem "D7"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 7
Beep 2093#, txtDuration.Text
Text1.Text = Text1.Text & "96 "
List1.AddItem 2093#
List2.AddItem txtDuration.Text
List3.AddItem "C7"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 8
Beep 1975.5, txtDuration.Text
Text1.Text = Text1.Text & "95 "
List1.AddItem 1975.5
List2.AddItem txtDuration.Text
List3.AddItem "B6"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 9
Beep 1760#, txtDuration.Text
Text1.Text = Text1.Text & "93 "
List1.AddItem 1760#
List2.AddItem txtDuration.Text
List3.AddItem "A6"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 10
Beep 1568#, txtDuration.Text
Text1.Text = Text1.Text & "91 "
List1.AddItem 1568#
List2.AddItem txtDuration.Text
List3.AddItem "G6"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 11
Beep 1396.9, txtDuration.Text
Text1.Text = Text1.Text & "89 "
List1.AddItem 1369.9
List2.AddItem txtDuration.Text
List3.AddItem "F6"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 12
Beep 1318.5, txtDuration.Text
Text1.Text = Text1.Text & "88 "
List1.AddItem 1318.5
List2.AddItem txtDuration.Text
List3.AddItem "E6"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 13
Beep 1174.7, txtDuration.Text
Text1.Text = Text1.Text & "86 "
List1.AddItem 1174.7
List2.AddItem txtDuration.Text
List3.AddItem "D6"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 14
Beep 1046.5, txtDuration.Text
Text1.Text = Text1.Text & "84 "
List1.AddItem 1046.5
List2.AddItem txtDuration.Text
List3.AddItem "C6"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 15
Beep 987.77, txtDuration.Text
Text1.Text = Text1.Text & "84 "
List1.AddItem 987.77
List2.AddItem txtDuration.Text
List3.AddItem "B5"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 16
Beep 880#, txtDuration.Text
Text1.Text = Text1.Text & "81 "
List1.AddItem 880#
List2.AddItem txtDuration.Text
List3.AddItem "A5"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 17
Beep 783.99, txtDuration.Text
Text1.Text = Text1.Text & "79 "
List1.AddItem 783.99
List2.AddItem txtDuration.Text
List3.AddItem "G5"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 18
Beep 698.46, txtDuration.Text
Text1.Text = Text1.Text & "77 "
List1.AddItem 689.46
List2.AddItem txtDuration.Text
List3.AddItem "F5"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 19
Beep 659.26, txtDuration.Text
Text1.Text = Text1.Text & "76 "
List1.AddItem 659.26
List2.AddItem txtDuration.Text
List3.AddItem "E5"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 20
Beep 587.33, txtDuration.Text
Text1.Text = Text1.Text & "74 "
List1.AddItem 587.33
List2.AddItem txtDuration.Text
List3.AddItem "D5"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 21
Beep 523.25, txtDuration.Text
Text1.Text = Text1.Text & "72 "
List1.AddItem 523.25
List2.AddItem txtDuration.Text
List3.AddItem "C5"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 22
Beep 493.38, txtDuration.Text
Text1.Text = Text1.Text & "71 "
List1.AddItem 493.38
List2.AddItem txtDuration.Text
List3.AddItem "B4"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 23
Beep 440#, txtDuration.Text
Text1.Text = Text1.Text & "69 "
List1.AddItem 440#
List2.AddItem txtDuration.Text
List3.AddItem "A4"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 24
Beep 392#, txtDuration.Text
Text1.Text = Text1.Text & "67 "
List1.AddItem 392#
List2.AddItem txtDuration.Text
List3.AddItem "G4"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 25
Beep 349.23, txtDuration.Text
Text1.Text = Text1.Text & "65 "
List1.AddItem 349.23
List2.AddItem txtDuration.Text
List3.AddItem "F4"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 26
Beep 329.63, txtDuration.Text
Text1.Text = Text1.Text & "64 "
List1.AddItem 329.63
List2.AddItem txtDuration.Text
List3.AddItem "E4"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 27
Beep 293.67, txtDuration.Text
Text1.Text = Text1.Text & "62 "
List1.AddItem 293.67
List2.AddItem txtDuration.Text
List3.AddItem "D4"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 28 ' middle c
Beep 261.6, txtDuration.Text
Text1.Text = Text1.Text & "60 "
List1.AddItem 261.6
List2.AddItem txtDuration.Text
List3.AddItem "C4"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 29
Beep 246.94, txtDuration.Text
Text1.Text = Text1.Text & "59 "
List1.AddItem 246.94
List2.AddItem txtDuration.Text
List3.AddItem "B3"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 30
Beep 220#, txtDuration.Text
Text1.Text = Text1.Text & "57 "
List1.AddItem 220#
List2.AddItem txtDuration.Text
List3.AddItem "A3"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 31
Beep 196#, txtDuration.Text
Text1.Text = Text1.Text & "55 "
List1.AddItem 196#
List2.AddItem txtDuration.Text
List3.AddItem "G3"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 32
Beep 174.61, txtDuration.Text
Text1.Text = Text1.Text & "53 "
List1.AddItem 174.61
List2.AddItem txtDuration.Text
List3.AddItem "F3"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 33
Beep 164.81, txtDuration.Text
Text1.Text = Text1.Text & "52 "
List1.AddItem 164.81
List2.AddItem txtDuration.Text
List3.AddItem "E3"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 34
Beep 146.83, txtDuration.Text
Text1.Text = Text1.Text & "50 "
List1.AddItem 146.83
List2.AddItem txtDuration.Text
List3.AddItem "D3"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 35
Beep 130.81, txtDuration.Text
Text1.Text = Text1.Text & "48 "
List1.AddItem 130.81
List2.AddItem txtDuration.Text
List3.AddItem "C3"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 36
Beep 123.47, txtDuration.Text
Text1.Text = Text1.Text & "47 "
List1.AddItem 123.47
List2.AddItem txtDuration.Text
List3.AddItem "B2"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 37
Beep 110#, txtDuration.Text
Text1.Text = Text1.Text & "45 "
List1.AddItem 110#
List2.AddItem txtDuration.Text
List3.AddItem "A2"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 38
Beep 97.999, txtDuration.Text
Text1.Text = Text1.Text & "43 "
List1.AddItem 97.999
List2.AddItem txtDuration.Text
List3.AddItem "G2"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 39
Beep 87.307, txtDuration.Text
Text1.Text = Text1.Text & "41 "
List1.AddItem 87.307
List2.AddItem txtDuration.Text
List3.AddItem "F2"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 40
Beep 82.407, txtDuration.Text
Text1.Text = Text1.Text & "40 "
List1.AddItem 82.407
List2.AddItem txtDuration.Text
List3.AddItem "E2"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 41
Beep 73.416, txtDuration.Text
Text1.Text = Text1.Text & "38 "
List1.AddItem 73.416
List2.AddItem txtDuration.Text
List3.AddItem "D2"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 42
Beep 65.406, txtDuration.Text
Text1.Text = Text1.Text & "36 "
List1.AddItem 65.407
List2.AddItem txtDuration.Text
List3.AddItem "C2"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 43
Beep 61.735, txtDuration.Text
Text1.Text = Text1.Text & "35 "
List1.AddItem 61.735
List2.AddItem txtDuration.Text
List3.AddItem "B1"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 44
Beep 55#, txtDuration.Text
Text1.Text = Text1.Text & "33 "
List1.AddItem 55#
List2.AddItem txtDuration.Text
List3.AddItem "A1"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 45
Beep 48.999, txtDuration.Text
Text1.Text = Text1.Text & "31 "
List1.AddItem 48.999
List2.AddItem txtDuration.Text
List3.AddItem "G1"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 46
Beep 43.654, txtDuration.Text
Text1.Text = Text1.Text & "29 "
List1.AddItem 43.654
List2.AddItem txtDuration.Text
List3.AddItem "F1"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 47
Beep 41.203, txtDuration.Text
Text1.Text = Text1.Text & "28 "
List1.AddItem 41.203
List2.AddItem txtDuration.Text
List3.AddItem "E1"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 48
Beep 36.708, txtDuration.Text
Text1.Text = Text1.Text & "26 "
List1.AddItem 36.708
List2.AddItem txtDuration.Text
List3.AddItem "D1"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 49
Beep 0, txtDuration.Text
Text1.Text = Text1.Text & "00 "
List1.AddItem 0
List2.AddItem txtDuration.Text
List3.AddItem "OOO"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
End Select
End Sub

Private Sub Picture1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************
'         set the treble cursor            *
'*******************************************
For Index = 0 To 48
Picture1(Index).MousePointer = 99
Picture1(Index).MouseIcon = LoadResPicture(102, vbResCursor)
Next Index
Picture1(49).MousePointer = 99
Picture1(49).MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub Picture2_Click(Index As Integer)
'*******************************************
'      play the sharp keys (black)         *
'   adds the note info to the listboxes    *
'     totals the number of notes and       *
'       totals the compostion time         *
'*******************************************
Select Case Index
Case 0
Beep 3729.3, txtDuration.Text
Text1.Text = Text1.Text & "106 "
List1.AddItem 3729.3
List2.AddItem txtDuration.Text
List3.AddItem "A7#"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 1
Beep 3322.4, txtDuration.Text
Text1.Text = Text1.Text & "104 "
List1.AddItem 3322.4
List2.AddItem txtDuration.Text
List3.AddItem "G7#"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 2
Beep 2960#, txtDuration.Text
Text1.Text = Text1.Text & "102 "
List1.AddItem 2960#
List2.AddItem txtDuration.Text
List3.AddItem "F7#"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 3
Beep 2498#, txtDuration.Text
Text1.Text = Text1.Text & "99 "
List1.AddItem 2498#
List2.AddItem txtDuration.Text
List3.AddItem "D7#"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 4
Beep 2217.5, txtDuration.Text
Text1.Text = Text1.Text & "97 "
List1.AddItem 2217.5
List2.AddItem txtDuration.Text
List3.AddItem "C7#"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 5
Beep 1864.7, txtDuration.Text
Text1.Text = Text1.Text & "94 "
List1.AddItem 1864.7
List2.AddItem txtDuration.Text
List3.AddItem "A6#"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 6
Beep 1661.2, txtDuration.Text
Text1.Text = Text1.Text & "92 "
List1.AddItem 1661.2
List2.AddItem txtDuration.Text
List3.AddItem "G6#"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 7
Beep 1480#, txtDuration.Text
Text1.Text = Text1.Text & "90 "
List1.AddItem 1480
List2.AddItem txtDuration.Text
List3.AddItem "F6#"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 8
Beep 1244.5, txtDuration.Text
Text1.Text = Text1.Text & "87 "
List1.AddItem 1244.5
List2.AddItem txtDuration.Text
List3.AddItem "D6#"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 9
Beep 1108.7, txtDuration.Text
Text1.Text = Text1.Text & "85 "
List1.AddItem 1108.7
List2.AddItem txtDuration.Text
List3.AddItem "C6#"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 10
Beep 932.33, txtDuration.Text
Text1.Text = Text1.Text & "82 "
List1.AddItem 932.33
List2.AddItem txtDuration.Text
List3.AddItem "A5#"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 11
Beep 830.61, txtDuration.Text
Text1.Text = Text1.Text & "80 "
List1.AddItem 830.61
List2.AddItem txtDuration.Text
List3.AddItem "G5#"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 12
Beep 739.99, txtDuration.Text
Text1.Text = Text1.Text & "78 "
List1.AddItem 739.99
List2.AddItem txtDuration.Text
List3.AddItem "F5#"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 13
Beep 622.25, txtDuration.Text
Text1.Text = Text1.Text & "75 "
List1.AddItem 622.25
List2.AddItem txtDuration.Text
List3.AddItem "D5#"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 14
Beep 554.37, txtDuration.Text
Text1.Text = Text1.Text & "73 "
List1.AddItem 554.37
List2.AddItem txtDuration.Text
List3.AddItem "C5#"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 15
Beep 466.16, txtDuration.Text
Text1.Text = Text1.Text & "70 "
List1.AddItem 466.16
List2.AddItem txtDuration.Text
List3.AddItem "A4#"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 16
Beep 415.3, txtDuration.Text
Text1.Text = Text1.Text & "68 "
List1.AddItem 415.3
List2.AddItem txtDuration.Text
List3.AddItem "G4#"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 17
Beep 369.99, txtDuration.Text
Text1.Text = Text1.Text & "66 "
List1.AddItem 369.99
List2.AddItem txtDuration.Text
List3.AddItem "F4#"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 18
Beep 311.13, txtDuration.Text
Text1.Text = Text1.Text & "63 "
List1.AddItem 311.13
List2.AddItem txtDuration.Text
List3.AddItem "D4#"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 19
Beep 277.18, txtDuration.Text
Text1.Text = Text1.Text & "61 "
List1.AddItem 277.18
List2.AddItem txtDuration.Text
List3.AddItem "C4#"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 20
Beep 233.08, txtDuration.Text
Text1.Text = Text1.Text & "58 "
List1.AddItem 233.08
List2.AddItem txtDuration.Text
List3.AddItem "A3#"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 21
Beep 207.65, txtDuration.Text
Text1.Text = Text1.Text & "56 "
List1.AddItem 207.65
List2.AddItem txtDuration.Text
List3.AddItem "G3#"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 22
Beep 185#, txtDuration.Text
Text1.Text = Text1.Text & "54 "
List1.AddItem 185#
List2.AddItem txtDuration.Text
List3.AddItem "F3#"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 23
Beep 155.56, txtDuration.Text
Text1.Text = Text1.Text & "51 "
List1.AddItem 155.56
List2.AddItem txtDuration.Text
List3.AddItem "D3#"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 24
Beep 138.59, txtDuration.Text
Text1.Text = Text1.Text & "49 "
List1.AddItem 138.59
List2.AddItem txtDuration.Text
List3.AddItem "C3#"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 25
Beep 116.54, txtDuration.Text
Text1.Text = Text1.Text & "46 "
List1.AddItem 116.54
List2.AddItem txtDuration.Text
List3.AddItem "A2#"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 26
Beep 103.83, txtDuration.Text
Text1.Text = Text1.Text & "44 "
List1.AddItem 103.83
List2.AddItem txtDuration.Text
List3.AddItem "G2#"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 27
Beep 92.499, txtDuration.Text
Text1.Text = Text1.Text & "42 "
List1.AddItem 92.499
List2.AddItem txtDuration.Text
List3.AddItem "F2#"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 28
Beep 77.782, txtDuration.Text
Text1.Text = Text1.Text & "39 "
List1.AddItem 77.782
List2.AddItem txtDuration.Text
List3.AddItem "D2#"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 29
Beep 69.296, txtDuration.Text
Text1.Text = Text1.Text & "37 "
List1.AddItem 69.296
List2.AddItem txtDuration.Text
List3.AddItem "C2#"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 30
Beep 58.27, txtDuration.Text
Text1.Text = Text1.Text & "34 "
List1.AddItem 58.27
List2.AddItem txtDuration.Text
List3.AddItem "A1#"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 31
Beep 51.913, txtDuration.Text
Text1.Text = Text1.Text & "31 "
List1.AddItem 51.913
List2.AddItem txtDuration.Text
List3.AddItem "G1#"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 32
Beep 46.249, txtDuration.Text
Text1.Text = Text1.Text & "30 "
List1.AddItem 46.249
List2.AddItem txtDuration.Text
List3.AddItem "F1#"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
Case 33
Beep 38.891, txtDuration.Text
Text1.Text = Text1.Text & "27 "
List1.AddItem 38.891
List2.AddItem txtDuration.Text
List3.AddItem "D1#"
txtPlayLister.Text = List1.ListCount
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
End Select
End Sub
Private Sub Picture2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************
'           set the bass cursor            *
'*******************************************
For Index = 0 To 33
Picture2(Index).MousePointer = 99
Picture2(Index).MouseIcon = LoadResPicture(105, vbResCursor)
Next Index
End Sub
Private Sub bkOpenList()
'*******************************************
'*        loads the tune list file         *
'*******************************************
Dim Filer As Integer
Dim lne$
Filer = FreeFile
Open App.Path & "\tunes.bkt" For Input As #Filer
Do While Not EOF(Filer)
Line Input #Filer, lne$
cmbTunelister.AddItem lne$
Loop
Close #Filer
End Sub
Private Sub bkSaveList()
'*******************************************
'*     saves the tune to the list file     *
'*******************************************
Dim Filer As Integer
Dim Msa As String
Filer = FreeFile
Msa = txtTuneSave.Text
Open App.Path & "\tunes.bkt" For Append As #Filer
Print #Filer, Msa
Close #Filer
cmbTunelister.AddItem txtTuneSave.Text
End Sub
Private Sub bkSaveFreq()
'*******************************************
'*  saves the frequency to the *.bkf file  *
'*******************************************
Dim Index As Integer
Dim Filer As Integer

Filer = FreeFile

Open App.Path & "\" & txtTuneSave.Text & ".bkf" For Append As #Filer

For Index = 0 To txtPlayLister.Text - 1
        Print #Filer, List1.List(Index)
    Next Index

    Close #Filer
End Sub
Private Sub bkSaveDur()
'*******************************************
'*  saves the duration to the *.bkd file  *
'*******************************************
Dim Index As Integer
Dim Filer As Integer

Filer = FreeFile

Open App.Path & "\" & txtTuneSave.Text & ".bkd" For Append As #Filer

For Index = 0 To txtPlayLister.Text - 1
        Print #Filer, List2.List(Index)
    Next Index

    Close #Filer
End Sub
Private Sub bkOpenFreq()
'*******************************************
'* loads the frequency from the *.bkf file *
'*******************************************
Dim Filer As Integer
Dim lne$
Filer = FreeFile
Open App.Path & "\" & cmbTunelister.Text & ".bkf" For Input As #Filer
Do While Not EOF(Filer)
Line Input #Filer, lne$
List1.AddItem lne$
Loop
Close #Filer
End Sub
Private Sub bkOpenDur()
'*******************************************
'* loads the duration from the *.bkd file *
'*******************************************
Dim Filer As Integer
Dim lne$
Filer = FreeFile
Open App.Path & "\" & cmbTunelister.Text & ".bkd" For Input As #Filer
Do While Not EOF(Filer)
Line Input #Filer, lne$
List2.AddItem lne$
Loop
Close #Filer
End Sub
Private Sub bkRptTime()
'**************************
'*  totals the tune time  *
'**************************
txtRepeatTime.Text = Val(txtRepeatTime.Text) + Val(txtDuration.Text)
End Sub

Private Sub LblLink_Click()
'*******************************************
' direct to user to website & open browser *
'*******************************************
   Dim lWindow As Long
    Call ShellExecute(lWindow, "open", "http://www.sds-software-maker.com/index.html", vbNullString, vbNullString, 5)
    End Sub
Private Sub LblLink_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'**********************************************************
'  mouse over effects hand cursor / bold / colour change  *
'**********************************************************
    LblLink.FontBold = True
    LblLink.FontUnderline = True
    LblLink.ForeColor = vbRed
    Me.MousePointer = 99
    Me.MouseIcon = LoadResPicture(101, vbResCursor)
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************
'  return label to original state on exit  *
'*******************************************
    LblLink.FontBold = False
    LblLink.FontUnderline = False
    LblLink.ForeColor = &H80FF&
    Me.MousePointer = 0
End Sub
Private Sub bkOpenNote()
'*******************************************
'*   loads the note from the *.bkn file   *
'*******************************************
Dim Filer As Integer
Dim lne$
Filer = FreeFile
Open App.Path & "\" & cmbTunelister.Text & ".bkn" For Input As #Filer
Do While Not EOF(Filer)
Line Input #Filer, lne$
List3.AddItem lne$
Loop
Close #Filer
End Sub
Private Sub bkSaveNote()
'*******************************************
'*    saves the note to the *.bkn file     *
'*******************************************
Dim Index As Integer
Dim Filer As Integer

Filer = FreeFile

Open App.Path & "\" & txtTuneSave.Text & ".bkn" For Append As #Filer

For Index = 0 To txtPlayLister.Text - 1
        Print #Filer, List3.List(Index)
    Next Index

    Close #Filer
End Sub
Private Sub List1_Scroll()
'******************************
'  synchs the list scrolling  *
'******************************
Call SendMessage(List2.hwnd, LB_SETTOPINDEX, SendMessage(List1.hwnd, LB_GETTOPINDEX, 0, 0), 0)
Call SendMessage(List3.hwnd, LB_SETTOPINDEX, SendMessage(List1.hwnd, LB_GETTOPINDEX, 0, 0), 0)

End Sub
Private Sub List2_Scroll()
'******************************
'  synchs the list scrolling  *
'******************************
Call SendMessage(List1.hwnd, LB_SETTOPINDEX, SendMessage(List2.hwnd, LB_GETTOPINDEX, 0, 0), 0)
Call SendMessage(List2.hwnd, LB_SETTOPINDEX, SendMessage(List1.hwnd, LB_GETTOPINDEX, 0, 0), 0)
Call SendMessage(List3.hwnd, LB_SETTOPINDEX, SendMessage(List2.hwnd, LB_GETTOPINDEX, 0, 0), 0)
End Sub
Private Sub List3_Scroll()
'******************************
'  synchs the list scrolling  *
'******************************
Call SendMessage(List1.hwnd, LB_SETTOPINDEX, SendMessage(List3.hwnd, LB_GETTOPINDEX, 0, 0), 0)
Call SendMessage(List2.hwnd, LB_SETTOPINDEX, SendMessage(List3.hwnd, LB_GETTOPINDEX, 0, 0), 0)
Call SendMessage(List3.hwnd, LB_SETTOPINDEX, SendMessage(List1.hwnd, LB_GETTOPINDEX, 0, 0), 0)
End Sub
