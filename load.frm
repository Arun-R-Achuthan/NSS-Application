VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form load 
   Caption         =   "Loading"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13575
   LinkTopic       =   "Form1"
   Picture         =   "load.frx":0000
   ScaleHeight     =   10500
   ScaleMode       =   0  'User
   ScaleWidth      =   1.81658e5
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   45
      Left            =   2520
      Top             =   7320
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   8400
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label p1 
      BackStyle       =   0  'Transparent
      Caption         =   "www.nss_bpc.com"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label wait 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   7320
      TabIndex        =   1
      Top             =   5760
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Shape s1 
      BorderColor     =   &H80000004&
      Height          =   3615
      Left            =   2760
      Top             =   1440
      Width           =   11655
   End
   Begin VB.Image Image1 
      Height          =   3405
      Left            =   2880
      Picture         =   "load.frx":1323B
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   11430
   End
End
Attribute VB_Name = "load"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
If pb.Value = 100 Then
pb.Value = 0
End If
nos = 0

End Sub

Private Sub Timer1_Timer()
If pb.Value < 100 Then

pb.Value = (pb.Value) + 1
If pb.Value Mod 2 = 0 Then
 wait.Visible = True
Else
wait.Visible = False
End If
End If
If pb.Value = 100 Then
wait.Caption = done
board.Show
Unload Me
End If
End Sub
