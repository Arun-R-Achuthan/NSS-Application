VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form camp 
   Caption         =   "Camp"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "hj.frx":0000
   ScaleHeight     =   10500
   ScaleWidth      =   18165
   WindowState     =   2  'Maximized
   Begin VB.TextBox cdetails 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   2775
      Left            =   4680
      MaxLength       =   300
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   6120
      Width           =   3975
   End
   Begin VB.TextBox cmax 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   5400
      MaxLength       =   2
      TabIndex        =   11
      Top             =   5280
      Width           =   855
   End
   Begin MSComCtl2.DTPicker cto 
      Height          =   375
      Left            =   4680
      TabIndex        =   10
      Top             =   4560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   16580609
      CurrentDate     =   42272
   End
   Begin MSComCtl2.DTPicker cfrom 
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   3840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   16580609
      CurrentDate     =   42272
   End
   Begin VB.TextBox clocation 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      MaxLength       =   30
      TabIndex        =   8
      Top             =   3120
      Width           =   3855
   End
   Begin VB.TextBox cname 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      MaxLength       =   20
      TabIndex        =   7
      Top             =   2400
      Width           =   3855
   End
   Begin MSFlexGridLib.MSFlexGrid flex 
      Height          =   7095
      Left            =   9240
      TabIndex        =   0
      Top             =   1920
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   12515
      _Version        =   393216
      Rows            =   1
      Cols            =   6
      Appearance      =   0
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   12120
      Top             =   9360
      Width           =   975
   End
   Begin VB.Label OK 
      BackStyle       =   0  'Transparent
      Caption         =   "  EXIT"
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
      TabIndex        =   19
      Top             =   9480
      Width           =   975
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   10920
      Top             =   9360
      Width           =   1095
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   9600
      Top             =   9360
      Width           =   1215
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   8280
      Top             =   9360
      Width           =   1215
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   7080
      Top             =   9360
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   5880
      Top             =   9360
      Width           =   1095
   End
   Begin VB.Label save 
      BackStyle       =   0  'Transparent
      Caption         =   "  SAVE"
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
      Left            =   10920
      TabIndex        =   18
      Top             =   9480
      Width           =   1095
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
      Left            =   9600
      TabIndex        =   17
      Top             =   9480
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
      Left            =   8280
      TabIndex        =   16
      Top             =   9480
      Width           =   1215
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
      Left            =   7080
      TabIndex        =   15
      Top             =   9480
      Width           =   1095
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
      Left            =   5880
      TabIndex        =   14
      Top             =   9480
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Details"
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
      Left            =   3480
      TabIndex        =   12
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Max Participants"
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
      Left            =   3480
      TabIndex        =   6
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
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
      Left            =   3480
      TabIndex        =   5
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "From"
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
      Left            =   3480
      TabIndex        =   4
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label Label3 
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
      Left            =   3480
      TabIndex        =   3
      Top             =   3120
      Width           =   975
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
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Camp Details"
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
      Left            =   8280
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      Height          =   7095
      Left            =   3120
      Top             =   1920
      Width           =   5895
   End
End
Attribute VB_Name = "camp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim sid As Integer
Dim flag As Integer
Public Function active()
cname.Enabled = True
cfrom.Enabled = True
cto.Enabled = True
cmax.Enabled = True
cdetails.Enabled = True
clocation.Enabled = True
End Function
Public Function dactive()
cname.Enabled = False
cfrom.Enabled = False
cto.Enabled = False
cmax.Enabled = False
cdetails.Enabled = False
clocation.Enabled = False
End Function
Public Function clearit()
cname.Text = ""
cdetails.Text = ""
cmax.Text = ""
clocation.Text = ""
End Function
Public Function fillgrid()
flex.TextMatrix(0, 0) = "No"
flex.TextMatrix(0, 1) = "Name"
flex.TextMatrix(0, 2) = "From"
flex.TextMatrix(0, 3) = "To"
flex.TextMatrix(0, 4) = "Max Participants"
flex.TextMatrix(0, 5) = "Location"
recordcheck
rs.Open ("select cno,cname,cfrom,cto,cmax,clocation from camp"), con, adOpenDynamic, adLockOptimistic
flex.Rows = 1
i = 1
While (rs.EOF = False)
flex.Rows = flex.Rows + 1
flex.TextMatrix(i, 0) = rs.Fields(0)
flex.TextMatrix(i, 1) = rs.Fields(1)
flex.TextMatrix(i, 2) = rs.Fields(2)
flex.TextMatrix(i, 3) = rs.Fields(3)
flex.TextMatrix(i, 4) = rs.Fields(4)
flex.TextMatrix(i, 5) = rs.Fields(5)
rs.MoveNext
i = i + 1
Wend
End Function


Private Sub add_Click()
add.Enabled = False
save.Enabled = True
flag = 1
active
End Sub
Private Sub clear_Click()
clearit
dactive
End Sub
Private Sub clocation_LostFocus()
If Not ValidName(clocation.Text) Then
 MsgBox ("Enter a valid location")
 clocation.SetFocus
End If
End Sub
Private Sub cmax_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
Else
KeyAscii = 0
End If
End Sub

Private Sub cname_LostFocus()
If Not ValidName(cname.Text) Then
 MsgBox ("Enter a valid name")
 cname.SetFocus
End If
End Sub




Private Sub cto_LostFocus()
If cto.Value < cfrom.Value Then
  MsgBox ("Enter correct from and to dates")
  cto.SetFocus
End If
End Sub

Private Sub delete_Click()
recordcheck
k = MsgBox("are you sure you want to delete", vbOKCancel)
If k = 1 Then
con.Execute ("delete  from camp where cno='" & sid & "'")
MsgBox ("Succesfully deleted")
clearit
dactive
fillgrid
recordcheck
Else
activity.Show
clearit
dactive
End If
delete.Enabled = False
edit.Enabled = False
add.Enabled = True
End Sub

Private Sub edit_Click()
edit.Enabled = False
save.Enabled = True
delete.Enabled = False
flag = 2
active
End Sub

Private Sub flex_Click()
sid = flex.TextMatrix(flex.Row, 0)
recordcheck
rs.Open ("select * from camp where cno='" & sid & "'"), con, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
cname.Text = rs.Fields(1)
cfrom.Value = rs.Fields(2)
cto.Value = rs.Fields(3)
cmax.Text = rs.Fields(4)
cdetails.Text = rs.Fields(5)
clocation.Text = rs.Fields(6)
End If
fillgrid
If nos = 1 Then
save.Enabled = False
add.Enabled = False
edit.Enabled = True
delete.Enabled = True
clear.Enabled = False
End If
End Sub

Private Sub Form_Load()
If nos = 2 Then
add.Enabled = False
edit.Enabled = False
clear.Enabled = False
save.Enabled = False
delete.Enabled = False
End If
If nos = 1 Then
add.Enabled = True
edit.Enabled = False
clear.Enabled = True
save.Enabled = False
delete.Enabled = False
End If
dactive
connection
recordcheck
rs.Open ("select * from camp "), con, adOpenDynamic, adLockOptimistic
i = 1
recordcheck
fillgrid
End Sub

Private Sub ok_Click()
main.Show
Unload Me
End Sub

Private Sub save_Click()
If cname.Text = "" Or cfrom.Value = "" Or cto.Value = "" Or cmax.Text = "" Or clocation.Text = "" Or cdetails.Text = "" Then
 MsgBox ("Please enter all the fields")
Else
save.Enabled = False
recordcheck
If flag = 1 Then
con.Execute ("insert into camp values('" & UCase(cname.Text) & "','" & cfrom.Value & "','" & cto.Value & "','" & UCase(cmax.Text) & "','" & UCase(cdetails.Text) & "','" & UCase(clocation.Text) & "')")
MsgBox ("Succesfully saved")
fillgrid
dactive
clearit
End If
If flag = 2 Then
con.Execute ("update camp set cname='" & UCase(cname.Text) & "',cfrom='" & cfrom.Value & "',cto='" & cto.Value & "',cmax='" & UCase(cmax.Text) & "',cdetails='" & UCase(cdetails.Text) & "', clocation='" & UCase(clocation.Text) & "' where cno = '" & sid & "'")
MsgBox ("Successfuly edited")
fillgrid
recordcheck
dactive
clearit
End If
add.Enabled = True
edit.Enabled = False
delete.Enabled = False
End If
End Sub
