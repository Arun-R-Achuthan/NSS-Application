VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fundraise 
   Caption         =   "Fundraise"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13380
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "fundraise.frx":0000
   ScaleHeight     =   7725
   ScaleWidth      =   13380
   WindowState     =   2  'Maximized
   Begin VB.ComboBox fgender 
      Height          =   315
      ItemData        =   "fundraise.frx":1323B
      Left            =   5880
      List            =   "fundraise.frx":13245
      TabIndex        =   15
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox fage 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   5880
      MaxLength       =   3
      TabIndex        =   14
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox id 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   6480
      TabIndex        =   11
      Top             =   5040
      Width           =   2415
   End
   Begin VB.TextBox fdisease 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   5880
      MaxLength       =   20
      TabIndex        =   10
      Top             =   4440
      Width           =   3015
   End
   Begin VB.TextBox fcontact 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   5880
      TabIndex        =   9
      Top             =   3840
      Width           =   3015
   End
   Begin VB.TextBox faddress 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   5880
      MaxLength       =   40
      TabIndex        =   8
      Top             =   3240
      Width           =   3975
   End
   Begin VB.TextBox fname 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   5880
      MaxLength       =   25
      TabIndex        =   7
      Top             =   1560
      Width           =   3975
   End
   Begin MSFlexGridLib.MSFlexGrid flex 
      Height          =   4815
      Left            =   10440
      TabIndex        =   1
      Top             =   1320
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   8493
      _Version        =   393216
      Rows            =   1
      Cols            =   8
      Appearance      =   0
   End
   Begin VB.Shape s5 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   3000
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label clear 
      BackStyle       =   0  'Transparent
      Caption         =   " CLEAR"
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
      Left            =   3000
      TabIndex        =   21
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   3000
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label ok 
      BackStyle       =   0  'Transparent
      Caption         =   "    EXIT"
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
      Left            =   3000
      TabIndex        =   20
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Shape s1 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   3000
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label edit 
      BackStyle       =   0  'Transparent
      Caption         =   "  EDIT"
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
      Left            =   3000
      TabIndex        =   19
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Shape s4 
      BorderColor     =   &H8000000E&
      Height          =   495
      Left            =   3000
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Shape s2 
      BorderColor     =   &H8000000E&
      Height          =   495
      Left            =   3000
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Shape s3 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   3000
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Save 
      BackStyle       =   0  'Transparent
      Caption         =   "   SAVE"
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
      Left            =   3000
      TabIndex        =   18
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label delete 
      BackStyle       =   0  'Transparent
      Caption         =   " DELETE"
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
      Left            =   3000
      TabIndex        =   17
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label add 
      BackStyle       =   0  'Transparent
      Caption         =   "   ADD"
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
      Left            =   3000
      TabIndex        =   16
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
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
      Left            =   4680
      TabIndex        =   13
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Age"
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
      Left            =   4680
      TabIndex        =   12
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Volunteer ID"
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
      Left            =   4680
      TabIndex        =   6
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Disease"
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
      Left            =   4680
      TabIndex        =   5
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact"
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
      Height          =   255
      Left            =   4680
      TabIndex        =   4
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Left            =   4680
      TabIndex        =   3
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      Left            =   4680
      TabIndex        =   2
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Financial Support"
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
      Height          =   735
      Left            =   9480
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      Height          =   4815
      Left            =   4440
      Top             =   1320
      Width           =   5655
   End
End
Attribute VB_Name = "fundraise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim sid As Integer
Dim flag As Integer
Public Function clearit()
fname.Text = ""
faddress.Text = ""
fcontact.Text = ""
fgender.Text = ""
fage.Text = ""
fdisease.Text = ""
If nos = 2 Then
id.Text = min
End If
If nos = 1 Then
id.Text = ""
End If
End Function
Public Function active()
fname.Enabled = True
fage.Enabled = True
faddress.Enabled = True
fcontact.Enabled = True
fgender.Enabled = True
fdisease.Enabled = True
If nos = 2 Then
id.Text = min
End If
End Function
Public Function dactive()
fname.Enabled = False
fage.Enabled = False
faddress.Enabled = False
fcontact.Enabled = False
fgender.Enabled = False
fdisease.Enabled = False
id.Enabled = False
If nos = 2 Then
id.Text = min
End If
End Function
Public Function fillgrid()
flex.TextMatrix(0, 0) = "N0"
flex.TextMatrix(0, 1) = "Name"
flex.TextMatrix(0, 2) = "Disease"
flex.TextMatrix(0, 3) = "Contact"
recordcheck
rs.Open ("select fno,fname,fdisease,fcontact from fund"), con, adOpenDynamic, adLockOptimistic
flex.Rows = 1
i = 1
While (rs.EOF = False)
flex.Rows = flex.Rows + 1
flex.TextMatrix(i, 0) = rs.Fields(0)
flex.TextMatrix(i, 1) = rs.Fields(1)
flex.TextMatrix(i, 2) = rs.Fields(2)
flex.TextMatrix(i, 3) = rs.Fields(3)
rs.MoveNext
i = i + 1
Wend
End Function
Private Sub add_Click()
active
flag = 1
add.Enabled = False
save.Enabled = True
End Sub

Private Sub clear_Click()
clearit
End Sub
Private Sub delete_Click()
recordcheck
k = MsgBox("are you sure you want to delete", vbOKCancel)
If k = 1 Then
rs.Open ("select * from fund where fno ='" & sid & "'"), con, adOpenDynamic, adLockOptimistic
con.Execute ("delete from fund where fno ='" & sid & "'")
MsgBox ("Succesfully deleted")
dactive
fillgrid
recordcheck
Else
fundraise.Show
dactive
End If
delete.Enabled = False
edit.Enabled = False
If nos = 1 Then
 add.Enabled = True
End If
End Sub

Private Sub edit_Click()
active
flag = 2
save.Enabled = True
edit.Enabled = False
delete.Enabled = False
End Sub
Private Sub faddress_LostFocus()
If Not ValidName(faddress.Text) Then
 MsgBox ("Enter a valid name")
 faddress.SetFocus
End If
End Sub

Private Sub fage_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 And kesyascii = 8 Then
Else
KeyAscii = 0
End If
End Sub

Private Sub fage_LostFocus()
If fage.Text > 120 Then
 MsgBox ("Please enter a valid age")
 fage.SetFocus
End If
End Sub

Private Sub fcontact_LostFocus()
If Not ValidPhone(fcontact.Text) Then
 MsgBox ("Enter a valid number")
 fcontact.SetFocus
End If
End Sub
Private Sub fdisease_LostFocus()
If Not ValidName(fdisease.Text) Then
 MsgBox ("Enter a valid disease")
 fdisease.SetFocus
End If
End Sub

Private Sub flex_Click()
sid = flex.TextMatrix(flex.Row, 0)
recordcheck
rs.Open ("select * from fund where fno='" & sid & "'"), con, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
fname.Text = rs.Fields(1)
fdisease.Text = rs.Fields(2)
fcontact.Text = rs.Fields(3)
fgender.Text = rs.Fields(4)
fage.Text = rs.Fields(5)
faddress.Text = rs.Fields(6)
id.Text = rs.Fields(7)
End If
If nos = 1 Then
 edit.Enabled = True
 delete.Enabled = True
End If
dactive
End Sub



Private Sub fname_LostFocus()
If Not ValidName(fname.Text) Then
 MsgBox ("Enter a valid name")
 fname.SetFocus
End If
End Sub

Private Sub Form_Load()
dactive
If nos = 2 Then
id.Text = min
add.Enabled = True
flex.Visible = False
edit.Visible = False
delete.Visible = False
s1.Visible = False
s2.Visible = False
End If
If nos = 1 Then
add.Visible = False
save.Visible = True
edit.Visible = True
s3.Visible = False
s4.Visible = True
s5.Visible = True
save.Enabled = False
edit.Enabled = True
clear.Enabled = False
delete.Enabled = True
End If
connection
recordcheck
fillgrid
End Sub
Private Sub ok_Click()
Unload Me
main.Show
End Sub

Private Sub save_Click()
If fname.Text = "" Or fage.Text = "" Or fcontact.Text = "" Or id.Text = "" Or fgender.Text = "" Or faddress.Text = "" Or fcontact.Text = "" Or fdisease.Text = "" Then
 MsgBox ("Please enter all the fields")
Else
recordcheck
If flag = 1 Then
con.Execute ("insert into fund values('" & UCase(fname.Text) & "','" & UCase(fdisease.Text) & "','" & UCase(fcontact.Text) & "','" & UCase(fgender.Text) & "','" & fage.Text & "','" & UCase(faddress.Text) & "','" & min & "')")
MsgBox ("Succesfully saved")
fillgrid
clearit
End If
If flag = 2 Then
con.Execute ("update fund set fname='" & UCase(fname.Text) & "',fdisease='" & UCase(fdisease.Text) & "',fcontact='" & UCase(fcontact.Text) & "',fgender='" & UCase(fgender.Text) & "',fage='" & UCase(fage.Text) & "',faddress='" & UCase(faddress.Text) & "',id='" & UCase(id.Text) & "'where fno='" & sid & "' ")
MsgBox ("Successfuly edited")
fillgrid
recordcheck
clearit
End If
End If
End Sub
