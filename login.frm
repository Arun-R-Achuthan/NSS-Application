VERSION 5.00
Begin VB.Form login 
   Caption         =   "Login"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   Picture         =   "login.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.TextBox pass 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   6600
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "-password-"
      Top             =   6120
      Width           =   4575
   End
   Begin VB.TextBox user 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Height          =   495
      Left            =   6600
      TabIndex        =   0
      Text            =   "-username-"
      Top             =   5400
      Width           =   4575
   End
   Begin VB.Label tittle 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to NSS BPC College"
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
      Left            =   6360
      TabIndex        =   4
      Top             =   480
      Width           =   4695
   End
   Begin VB.Image Image1 
      Height          =   3300
      Left            =   7200
      Picture         =   "login.frx":1323B
      Top             =   1920
      Width           =   3390
   End
   Begin VB.Label forgot 
      BackStyle       =   0  'Transparent
      Caption         =   "Forgrot password?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   7560
      TabIndex        =   3
      Top             =   7680
      Width           =   2655
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   615
      Left            =   6600
      Top             =   6960
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "                        LOGIN"
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
      Height          =   495
      Left            =   6600
      TabIndex        =   2
      Top             =   7080
      Width           =   4575
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000E&
      Height          =   7335
      Left            =   4920
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   7335
   End
   Begin VB.Menu mnuexit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub forgot_Click()
fpassword.Show
Unload Me
End Sub

Private Sub Form_Load()
connection
recordcheck
nos = 0
End Sub
Private Sub Label1_Click()
recordcheck
If Trim(user.Text) = "admin" And Trim(pass.Text) = "admin" Then
nos = 1
Unload Me
main.Show
Else
recordcheck
rs.Open ("select username,password from enroll where username='" & user.Text & "'"), con, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
If Trim(user.Text) = rs.Fields(0) And Trim(pass.Text) = rs.Fields(1) Then
min = user.Text
nos = 2
Unload Me
main.Show

Else
MsgBox "Wrong username or password", vbCritical
user.Text = ""
pass.Text = ""
End If
Else
MsgBox "Wrong username or password", vbCritical
End If
End If
End Sub


Private Sub mnuexit_Click()
End
End Sub

Private Sub pass_Click()
If pass.Text = "-password-" Then
pass.Text = ""
End If

End Sub

Private Sub pass_LostFocus()
If pass.Text = "" Then
pass.Text = "-password-"
End If
End Sub


Private Sub user_Click()
If user.Text = "-username-" Then
user.Text = ""
End If

End Sub

Private Sub user_LostFocus()
If user.Text = "" Then
user.Text = "-username-"
End If
End Sub
