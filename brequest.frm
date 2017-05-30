VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form brequest 
   Caption         =   "Blood Request"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "brequest.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid flex 
      Height          =   6015
      Left            =   11760
      TabIndex        =   20
      Top             =   1080
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   10610
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      Appearance      =   0
   End
   Begin VB.TextBox id 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   7800
      TabIndex        =   19
      Top             =   6000
      Width           =   2295
   End
   Begin VB.TextBox bcontact 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   7320
      TabIndex        =   18
      Top             =   5400
      Width           =   2775
   End
   Begin MSComCtl2.DTPicker bdate 
      Height          =   375
      Left            =   7320
      TabIndex        =   17
      Top             =   4800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   68943873
      CurrentDate     =   42271
   End
   Begin VB.TextBox btime 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   10080
      MaxLength       =   7
      TabIndex        =   16
      Top             =   4200
      Width           =   1095
   End
   Begin VB.ComboBox bgroup 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "brequest.frx":1323B
      Left            =   7920
      List            =   "brequest.frx":13257
      TabIndex        =   15
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox blocation 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   7560
      MaxLength       =   30
      TabIndex        =   14
      Top             =   3480
      Width           =   3615
   End
   Begin VB.TextBox bhname 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   8040
      TabIndex        =   13
      Top             =   2760
      Width           =   3135
   End
   Begin VB.ComboBox bgender 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "brequest.frx":1327D
      Left            =   9480
      List            =   "brequest.frx":13287
      TabIndex        =   12
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox bage 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   7200
      MaxLength       =   3
      TabIndex        =   10
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox bname 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   7200
      MaxLength       =   25
      TabIndex        =   9
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   4320
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label12 
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
      Height          =   375
      Left            =   4320
      TabIndex        =   27
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Shape s2 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   4320
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label delete 
      BackStyle       =   0  'Transparent
      Caption         =   "  DELETE"
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
      Left            =   4320
      TabIndex        =   26
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Shape s5 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   4320
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Shape s4 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   4320
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Shape s1 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   4320
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Shape s3 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   4320
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label save 
      BackStyle       =   0  'Transparent
      Caption         =   "     SAVE"
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
      Left            =   4320
      TabIndex        =   25
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label clear 
      BackStyle       =   0  'Transparent
      Caption         =   "   CLEAR"
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
      Left            =   4320
      TabIndex        =   24
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label edit 
      BackStyle       =   0  'Transparent
      Caption         =   "   EDIT"
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
      Left            =   4320
      TabIndex        =   23
      Top             =   4680
      Width           =   1335
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
      Left            =   4320
      TabIndex        =   22
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Blood Request"
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
      Left            =   10800
      TabIndex        =   21
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label10 
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
      Left            =   8400
      TabIndex        =   11
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label9 
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
      Left            =   6240
      TabIndex        =   8
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
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
      Left            =   6240
      TabIndex        =   7
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label7 
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
      Height          =   375
      Left            =   6240
      TabIndex        =   6
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Blood Group"
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
      Left            =   6240
      TabIndex        =   5
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
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
      Left            =   6240
      TabIndex        =   4
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label Label4 
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
      Height          =   375
      Left            =   6240
      TabIndex        =   3
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
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
      Left            =   9360
      TabIndex        =   2
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Hospital Name"
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
      Left            =   6240
      TabIndex        =   1
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label1 
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
      Height          =   375
      Left            =   6240
      TabIndex        =   0
      Top             =   1560
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000004&
      Height          =   6015
      Left            =   6120
      Top             =   1080
      Width           =   5415
   End
End
Attribute VB_Name = "brequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim flag As Integer
Dim sid As Integer

Public Function fillgrid()
flex.TextMatrix(0, 0) = "N0"
flex.TextMatrix(0, 1) = "Name"
flex.TextMatrix(0, 2) = "Hosp Name"
flex.TextMatrix(0, 3) = "Contact"
flex.TextMatrix(0, 4) = "Date"
recordcheck
rs.Open ("select bno,bname,bhname,bcontact,bdate from blood"), con, adOpenDynamic, adLockOptimistic
flex.Rows = 1
i = 1
While (rs.EOF = False)
flex.Rows = flex.Rows + 1
flex.TextMatrix(i, 0) = rs.Fields(0)
flex.TextMatrix(i, 1) = rs.Fields(1)
flex.TextMatrix(i, 2) = rs.Fields(2)
flex.TextMatrix(i, 3) = rs.Fields(3)
flex.TextMatrix(i, 4) = rs.Fields(4)
rs.MoveNext
i = i + 1
Wend
End Function
Public Function clearit()
bname.Text = ""
bage.Text = ""
bhname.Text = ""
blocation.Text = ""
btime.Text = ""
bcontact.Text = ""
If nos = 2 Then
id.Text = min
End If
End Function
Public Function active()
bname.Enabled = True
bage.Enabled = True
bgender.Enabled = True
bhname.Enabled = True
blocation.Enabled = True
bgroup.Enabled = True
btime.Enabled = True
bdate.Enabled = True
bcontact.Enabled = True
End Function
Public Function dactive()
bname.Enabled = False
bage.Enabled = False
bgender.Enabled = False
bhname.Enabled = False
blocation.Enabled = False
bgroup.Enabled = False
btime.Enabled = False
bdate.Enabled = False
bcontact.Enabled = False
id.Enabled = False
End Function

Private Sub add_Click()
flag = 1
active
add.Enabled = False
save.Enabled = True
End Sub
Private Sub bage_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
Else
KeyAscii = 0
End If
End Sub

Private Sub bage_LostFocus()
If bage.Text > 120 Then
 MsgBox ("Please enter a valid age")
 bage.SetFocus
End If
End Sub

Private Sub bcontact_LostFocus()
If Not ValidPhone(bcontact.Text) Then
 MsgBox ("Enter a valid contact")
 bcontact.SetFocus
End If
End Sub

Private Sub bhname_LostFocus()
If Not ValidName(bhname.Text) Then
 MsgBox ("Enter a valid name")
 bhname.SetFocus
End If
End Sub



Private Sub blocation_LostFocus()
If Not ValidName(blocation.Text) Then
 MsgBox ("Please enter a valid location")
 blocation.SetFocus
End If
End Sub

Private Sub bname_LostFocus()
If Not ValidName(bname.Text) Then
 MsgBox ("Enter a valid name")
 bname.SetFocus
End If
End Sub



Private Sub btime_Click()
MsgBox ("Please enter in the format, eg. 9.00AM")
End Sub

Private Sub clear_Click()
clearit
dactive
End Sub

Private Sub delete_Click()
recordcheck
k = MsgBox("are you sure you want to delete", vbOKCancel)
If k = 1 Then
con.Execute ("delete from blood where bno ='" & sid & "' ")
MsgBox ("Succesfully deleted")
dactive
fillgrid
clearit
recordcheck
Else
brequest.Show
clearit
dactive
End If
delete.Enabled = False
End Sub

Private Sub edit_Click()
flag = 2
active
edit.Enabled = False
save.Enabled = True
delete.Enabled = False
End Sub

Private Sub flex_Click()
sid = flex.TextMatrix(flex.Row, 0)
recordcheck
rs.Open ("select * from blood where bno='" & sid & "'"), con, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
bname.Text = rs.Fields(1)
bage.Text = rs.Fields(2)
bgender.Text = rs.Fields(3)
bhname.Text = rs.Fields(4)
blocation.Text = rs.Fields(5)
bgroup.Text = rs.Fields(6)
btime.Text = rs.Fields(7)
bdate.Value = rs.Fields(8)
bcontact.Text = rs.Fields(9)
id.Text = rs.Fields(10)
End If
If nos = 1 Then
edit.Enabled = True
delete.Enabled = True
End If
dactive
End Sub

Private Sub Form_Load()
connection
recordcheck
fillgrid
If nos = 1 Then
add.Visible = False
clear.Visible = False
save.Visible = True
s3.Visible = False
s4.Visible = False
s5.Visible = True
save.Enabled = False
edit.Enabled = False
delete.Enabled = False
End If
If nos = 2 Then
id.Text = min
add.Enabled = True
add.Visible = True
flex.Visible = False
delete.Visible = False
edit.Visible = False
s1.Visible = False
s2.Visible = False
save.Enabled = False
End If
End Sub

Private Sub Label12_Click()
main.Show
Unload Me
End Sub

Private Sub save_Click()
recordcheck
If flag = 1 Then
con.Execute ("insert into blood values('" & UCase(bname.Text) & "','" & UCase(bage.Text) & "','" & UCase(bgender.Text) & "','" & UCase(bhname.Text) & "','" & UCase(blocation.Text) & "','" & UCase(bgroup.Text) & "','" & btime.Text & "','" & bdate.Value & "','" & bcontact.Text & "','" & id.Text & "')")
MsgBox ("Succesfully saved")
fillgrid
clearit
dactive
End If
If flag = 2 Then
con.Execute ("update blood set bname='" & UCase(bname.Text) & "',bage='" & UCase(bage.Text) & "',bgender='" & UCase(bgender.Text) & "',bhname='" & UCase(bhname.Text) & "',blocation='" & UCase(blocation.Text) & "',bgroup='" & UCase(bgroup.Text) & "',btime='" & UCase(btime.Text) & "',bdate='" & bdate.Value & "',bcontact='" & UCase(bcontact.Text) & "',id='" & min & "' where bno = '" & sid & "' ")
MsgBox ("Successfuly edited")
dactive
fillgrid
recordcheck
clearit
End If
save.Enabled = False
If nos = 2 Then
add.Enabled = True
End If
End Sub
