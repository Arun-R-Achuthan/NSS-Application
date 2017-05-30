VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form bloodlist 
   Caption         =   "Bloodlist"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "bloodlist.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.TextBox bgrps 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   6360
      TabIndex        =   19
      Top             =   4440
      Width           =   1455
   End
   Begin VB.ComboBox dids 
      Height          =   315
      Left            =   6360
      TabIndex        =   18
      Top             =   3000
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid flex 
      Height          =   5415
      Left            =   9960
      TabIndex        =   10
      Top             =   1800
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   9551
      _Version        =   393216
      Rows            =   1
      Cols            =   4
      Appearance      =   0
   End
   Begin MSComCtl2.DTPicker bdates 
      Height          =   375
      Left            =   6360
      TabIndex        =   9
      Top             =   5760
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   69206017
      CurrentDate     =   42280
   End
   Begin VB.TextBox hnames 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   6360
      MaxLength       =   25
      TabIndex        =   8
      Top             =   5040
      Width           =   3135
   End
   Begin VB.TextBox dnames 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   6360
      TabIndex        =   7
      Top             =   3720
      Width           =   3135
   End
   Begin VB.TextBox rnames 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   6360
      MaxLength       =   20
      TabIndex        =   6
      Top             =   2280
      Width           =   3135
   End
   Begin VB.Label xzcbasc 
      BackStyle       =   0  'Transparent
      Caption         =   "Blood List"
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
      Height          =   375
      Left            =   8640
      TabIndex        =   17
      Top             =   840
      Width           =   1695
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   13680
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label Label7 
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
      Left            =   13680
      TabIndex        =   16
      Top             =   7800
      Width           =   1215
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H8000000E&
      Height          =   495
      Left            =   12120
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H8000000E&
      Height          =   495
      Left            =   10680
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H8000000E&
      Height          =   495
      Left            =   9000
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H8000000E&
      Height          =   495
      Left            =   7560
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H8000000E&
      Height          =   495
      Left            =   6120
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Label save 
      BackStyle       =   0  'Transparent
      Caption         =   "    SAVE"
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
      Left            =   12120
      TabIndex        =   15
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Label clear 
      BackStyle       =   0  'Transparent
      Caption         =   "  CLEAR"
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
      Left            =   10680
      TabIndex        =   14
      Top             =   7800
      Width           =   1335
   End
   Begin VB.Label delete 
      BackStyle       =   0  'Transparent
      Caption         =   "   DELETE"
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
      Left            =   9000
      TabIndex        =   13
      Top             =   7800
      Width           =   1575
   End
   Begin VB.Label edit 
      BackStyle       =   0  'Transparent
      Caption         =   "    EDIT"
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
      Left            =   7560
      TabIndex        =   12
      Top             =   7800
      Width           =   1335
   End
   Begin VB.Label add 
      BackStyle       =   0  'Transparent
      Caption         =   "    ADD"
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
      Left            =   6120
      TabIndex        =   11
      Top             =   7800
      Width           =   1335
   End
   Begin VB.Label Label6 
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
      Left            =   4200
      TabIndex        =   5
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Hospital name"
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
      Left            =   4200
      TabIndex        =   4
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label4 
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
      Left            =   4200
      TabIndex        =   3
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Donar ID"
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
      Left            =   4200
      TabIndex        =   2
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Donar name"
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
      Left            =   4200
      TabIndex        =   1
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Receiver name"
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
      Left            =   4200
      TabIndex        =   0
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000004&
      Height          =   5415
      Left            =   3720
      Top             =   1800
      Width           =   5895
   End
End
Attribute VB_Name = "bloodlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag As Integer
Dim sid As Integer
Public Function fillgrid()
flex.TextMatrix(0, 0) = "No"
flex.TextMatrix(0, 1) = "Receiver name"
flex.TextMatrix(0, 2) = "Donar name"
flex.TextMatrix(0, 3) = "Bgroup"
recordcheck
rs.Open ("select anos,rname,dname,bgrp  from bloodlist"), con, adOpenDynamic, adLockOptimistic
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
Public Function clearit()
rnames.Text = ""
dnames.Text = ""
dids.Text = ""
bgrps.Text = ""
hnames.Text = ""
dactive
End Function
Public Function dactive()
rnames.Enabled = False
dnames.Enabled = False
dids.Enabled = False
bgrps.Enabled = False
hnames.Enabled = False
bdates.Enabled = False
End Function
Public Function active()
rnames.Enabled = True
dids.Enabled = True
hnames.Enabled = True
bdates.Enabled = True
End Function

Private Sub add_Click()
add.Enabled = False
save.Enabled = True
flag = 1
active
End Sub



Private Sub bgrps_KeyPress(KeyAscii As Integer)
If KeyAscii > 0 Then
 KeyAscii = 0
End If
End Sub

Private Sub delete_Click()
delete.Enabled = False
add.Enabled = True
edit.Enabled = False
recordcheck
k = MsgBox("are you sure you want to delete", vbOKCancel)
If k = 1 Then
con.Execute ("delete  from bloodlist where anos='" & sid & "'")
MsgBox ("Succesfully deleted")
clearit
dactive
fillgrid
recordcheck
Else
bloodlist.Show
clearit
dactive
End If
End Sub





Private Sub dids_Click()
dnames.Enabled = True
bgrps.Enabled = True
End Sub

Private Sub dids_LostFocus()
If dids.Text <> "" Then
 recordcheck
 rs.Open ("select ename from enroll where id = '" & dids.Text & "'"), con, adOpenDynamic, adLockOptimistic
 dnames.Text = rs.Fields(0)
End If
If dids.Text <> "" Then
 recordcheck
 rs.Open ("select ebgroup from enroll where id = '" & dids.Text & "'"), con, adOpenDynamic, adLockOptimistic
 bgrps.Text = rs.Fields(0)
End If

End Sub



Private Sub dnames_KeyPress(KeyAscii As Integer)
If KeyAscii > 0 Then
 KeyAscii = 0
End If
End Sub

Private Sub dnames_LostFocus()
If Not ValidName(dnames.Text) Then
 MsgBox ("Enter a valid donar name")
 dnames.SetFocus
End If
End Sub

Private Sub flex_Click()
sid = flex.TextMatrix(flex.Row, 0)
recordcheck
rs.Open ("select * from bloodlist where anos='" & sid & "'"), con, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
rnames.Text = rs.Fields(1)
dnames.Text = rs.Fields(2)
dids.Text = rs.Fields(5)
bgrps.Text = rs.Fields(3)
bdates.Value = rs.Fields(4)
hnames.Text = rs.Fields(6)
End If
dactive
save.Enabled = False
edit.Enabled = True
delete.Enabled = True
add.Enabled = False
clear.Enabled = False
End Sub

Private Sub Form_Load()
add.Enabled = True
edit.Enabled = False
save.Enabled = False
delete.Enabled = False
clear.Enabled = True
dactive
connection
recordcheck
fillgrid
recordcheck
rs.Open ("select id from enroll"), con, adOpenDynamic, adLockOptimistic
While rs.EOF = False
dids.AddItem (rs.Fields(0))
rs.MoveNext
Wend
End Sub
Private Sub hnames_LostFocus()
If Not ValidName(hnames.Text) Then
 MsgBox ("Enter a valid name")
 hnames.SetFocus
End If
End Sub

Private Sub Label7_Click()
Unload Me
main.Show
End Sub



Private Sub rnames_LostFocus()
If Not ValidName(rnames.Text) Then
 MsgBox ("Please enter a valid name")
 rnames.SetFocus
End If
End Sub

Private Sub save_Click()
recordcheck
If flag = 1 Then
con.Execute ("update enroll set elast ='" & bdates.Value & "' where id = '" & dids.Text & "' ")
con.Execute ("insert into bloodlist values('" & UCase(rnames.Text) & "','" & UCase(dnames.Text) & "','" & UCase(bgrps.Text) & "','" & UCase(bdates.Value) & "','" & UCase(dids.Text) & "','" & UCase(hnames.Text) & "')")
MsgBox ("Last blood donation of this volunteer is updated and Succesfully saved")

fillgrid
dactive
clearit
End If
If flag = 2 Then
con.Execute ("update bloodlist set rname='" & UCase(rnames.Text) & "',dname='" & UCase(dnames.Text) & "',bgrp='" & UCase(bgrps.Text) & "',dates='" & bdates.Value & "',did='" & UCase(dids.Text) & "', hname='" & UCase(hnames.Text) & "' where anos = '" & sid & "'")
MsgBox ("Successfuly edited")
fillgrid
recordcheck
dactive
clearit
End If
save.Enabled = False
add.Enabled = True
End Sub

Private Sub edit_Click()
edit.Enabled = False
save.Enabled = True
flag = 2
active
End Sub

Private Sub clear_Click()
clearit
End Sub
