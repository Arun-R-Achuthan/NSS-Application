VERSION 5.00
Begin VB.Form fpassword 
   Caption         =   "Forgot Password"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "fpassword.frx":0000
   ScaleHeight     =   10500
   ScaleWidth      =   18165
   WindowState     =   2  'Maximized
   Begin VB.TextBox equestion 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   7800
      TabIndex        =   8
      Top             =   4920
      Width           =   3135
   End
   Begin VB.TextBox eanswer 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   7800
      TabIndex        =   5
      Top             =   5640
      Width           =   3135
   End
   Begin VB.TextBox user 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   7800
      TabIndex        =   2
      Top             =   4080
      Width           =   3135
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   10440
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label ok 
      BackStyle       =   0  'Transparent
      Caption         =   "   OK"
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
      Left            =   10440
      TabIndex        =   10
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1830
      Left            =   7200
      Picture         =   "fpassword.frx":1323B
      Top             =   1920
      Width           =   2130
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   5640
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Label Proceed 
      BackStyle       =   0  'Transparent
      Caption         =   "PROCEED"
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
      Left            =   5760
      TabIndex        =   9
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   9120
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   7320
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Label cancel 
      BackStyle       =   0  'Transparent
      Caption         =   "CANCEL"
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
      Left            =   9120
      TabIndex        =   7
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label retrieve 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   " RETRIEVE"
      Enabled         =   0   'False
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
      Left            =   7320
      TabIndex        =   6
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Answer"
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
      Left            =   5640
      TabIndex        =   4
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Security Question"
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
      Left            =   5640
      TabIndex        =   3
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
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
      Left            =   5640
      TabIndex        =   1
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Forgot Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   6840
      TabIndex        =   0
      Top             =   720
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      Height          =   5895
      Left            =   5280
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   6855
   End
End
Attribute VB_Name = "fpassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancel_Click()
login.Show
Unload Me
End Sub



Private Sub eanswer_LostFocus()
If Not ValidName(eanswer.Text) Then
 MsgBox ("Answer not valid")
 eanswer.SetFocus
End If
End Sub

Private Sub Form_Load()
connection
recordcheck
End Sub

Private Sub ok_Click()
login.Show
Unload Me
End Sub

Private Sub Proceed_Click()
eanswer.Enabled = True
recordcheck
rs.Open ("select username,password,equestion,eanswer from enroll where username='" & user.Text & "'"), con, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
If user.Text = rs.Fields(0) Then
 equestion.Text = rs.Fields(2)
 eanswer.Enabled = True
 retrieve.Enabled = True
 End If
  Else
 MsgBox "invalid username", vbCritical
 user.Text = ""
 End If
End Sub

Private Sub retrieve_Click()
If LCase(eanswer.Text) = rs.Fields(3) Then
MsgBox "Your password is " & rs.Fields(1)
eanswer.Text = ""
equestion.Text = ""
Else
MsgBox "Wrong answer please try again  ", vbCritical
eanswer.Text = ""
equestion.Text = ""
user.Text = ""
End If
End Sub
