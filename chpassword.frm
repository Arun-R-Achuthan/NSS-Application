VERSION 5.00
Begin VB.Form chpassword 
   Caption         =   "Change Password"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   Picture         =   "chpassword.frx":0000
   ScaleHeight     =   10500
   ScaleWidth      =   18165
   WindowState     =   2  'Maximized
   Begin VB.TextBox repass 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   8880
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   6840
      Width           =   3375
   End
   Begin VB.TextBox newpass 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   8880
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   6240
      Width           =   3375
   End
   Begin VB.TextBox pass 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   8880
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   5640
      Width           =   3375
   End
   Begin VB.TextBox user 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   8880
      TabIndex        =   4
      Top             =   5040
      Width           =   3375
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   615
      Left            =   11040
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Label ok 
      BackStyle       =   0  'Transparent
      Caption         =   "     OK"
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
      Left            =   11040
      TabIndex        =   11
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1425
      Left            =   8040
      Picture         =   "chpassword.frx":1323B
      Top             =   3240
      Width           =   3180
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000009&
      Height          =   615
      Left            =   6120
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Label Proceed 
      BackStyle       =   0  'Transparent
      Caption         =   "   Proceed"
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
      Left            =   6120
      TabIndex        =   10
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      Height          =   615
      Left            =   9360
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   615
      Left            =   7800
      Top             =   7560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label cancel 
      BackStyle       =   0  'Transparent
      Caption         =   "     Cancel"
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
      Left            =   9360
      TabIndex        =   9
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Re-enter Password"
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
      Left            =   6480
      TabIndex        =   8
      Top             =   6840
      Width           =   2175
   End
   Begin VB.Label change 
      BackStyle       =   0  'Transparent
      Caption         =   "   Change"
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
      Left            =   7800
      TabIndex        =   3
      Top             =   7680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "New Password"
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
      Left            =   6480
      TabIndex        =   2
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Password"
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
      Left            =   6480
      TabIndex        =   1
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Label Label1 
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
      Left            =   6480
      TabIndex        =   0
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000004&
      Height          =   5895
      Left            =   5880
      Shape           =   4  'Rounded Rectangle
      Top             =   2760
      Width           =   6855
   End
End
Attribute VB_Name = "chpassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cancel_Click()
main.Show
Unload Me
End Sub

Private Sub change_Click()
If newpass.Text = repass.Text Then
recordcheck
rs.Open ("update enroll set password='" & newpass.Text & "' where username='" & user.Text & "'"), con, adOpenDynamic, adLockOptimistic
MsgBox "Password updated successfully"
Else
MsgBox "passwords do not match", vbCritical
End If
newpass.Text = ""
user.Text = ""
pass.Text = ""
repass.Text = ""
End Sub

Private Sub Form_Load()
connection
recordcheck
End Sub

Private Sub ok_Click()
main.Show
Unload Me
End Sub

Private Sub Proceed_Click()
Proceed.Enabled = False
recordcheck
rs.Open ("select username,password from enroll where username='" & user.Text & "'"), con, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
If (user.Text = rs.Fields(0)) Then
If (pass.Text = rs.Fields(1)) Then
MsgBox ("Enter new password")
newpass.Enabled = True
repass.Enabled = True
Shape2.Visible = True
change.Visible = True
End If
End If
Else
MsgBox "invalid username or password", vbCritical
user.Text = ""
pass.Text = ""
End If
End Sub
