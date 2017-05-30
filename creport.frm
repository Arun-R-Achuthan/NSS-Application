VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form creport 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "creport.frx":0000
   ScaleHeight     =   10500
   ScaleWidth      =   18165
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid flex 
      Height          =   6855
      Left            =   10080
      TabIndex        =   5
      Top             =   1440
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   12091
      _Version        =   393216
      Appearance      =   0
      GridLineWidth   =   3
   End
   Begin VB.TextBox creports 
      Appearance      =   0  'Flat
      Height          =   3495
      Left            =   5280
      MaxLength       =   400
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   3840
      Width           =   3375
   End
   Begin VB.ComboBox cname 
      Height          =   315
      Left            =   5280
      TabIndex        =   2
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   13320
      Top             =   8520
      Width           =   975
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   12120
      Top             =   8520
      Width           =   1095
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   11040
      Top             =   8520
      Width           =   975
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   9840
      Top             =   8520
      Width           =   1095
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   8640
      Top             =   8520
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   7440
      Top             =   8520
      Width           =   1095
   End
   Begin VB.Label ok 
      BackStyle       =   0  'Transparent
      Caption         =   "   EXIT"
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
      Left            =   13320
      TabIndex        =   11
      Top             =   8640
      Width           =   975
   End
   Begin VB.Label delete 
      BackStyle       =   0  'Transparent
      Caption         =   "DELETE"
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
      TabIndex        =   10
      Top             =   8640
      Width           =   1095
   End
   Begin VB.Label save 
      BackStyle       =   0  'Transparent
      Caption         =   " SAVE"
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
      Left            =   11040
      TabIndex        =   9
      Top             =   8640
      Width           =   975
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
      Left            =   9840
      TabIndex        =   8
      Top             =   8640
      Width           =   1095
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
      Left            =   8640
      TabIndex        =   7
      Top             =   8640
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
      Left            =   7440
      TabIndex        =   6
      Top             =   8640
      Width           =   1095
   End
   Begin VB.Label sajd 
      BackStyle       =   0  'Transparent
      Caption         =   "Camp Report"
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
      Height          =   615
      Left            =   8880
      TabIndex        =   4
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Report"
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
      Left            =   3360
      TabIndex        =   1
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Camp Name"
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
      Left            =   3360
      TabIndex        =   0
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      Height          =   6735
      Left            =   3000
      Top             =   1440
      Width           =   6495
   End
End
Attribute VB_Name = "creport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag As Integer
Dim sid As Integer
Dim i As Integer
Public Function listing()
recordcheck
rs.Open ("select * from camp "), con, adOpenDynamic, adLockOptimistic
While rs.EOF = False
cname.AddItem (rs.Fields(1))
rs.MoveNext
Wend
End Function
Public Function active()
cname.Enabled = True
creports.Enabled = True
End Function
Public Function dactive()
cname.Enabled = False
creports.Enabled = False
End Function
Public Function fillgrid()
flex.TextMatrix(0, 0) = "No"
flex.TextMatrix(0, 1) = "Name"
recordcheck
rs.Open ("select cno,cname from creport"), con, adOpenDynamic, adLockOptimistic
flex.Rows = 1
i = 1
While (rs.EOF = False)
flex.Rows = flex.Rows + 1
flex.TextMatrix(i, 0) = rs.Fields(0)
flex.TextMatrix(i, 1) = rs.Fields(1)
rs.MoveNext
i = i + 1
Wend
delete.Enabled = True
edit.Enabled = True
add.Enabled = False
End Function
Public Function clearit()
cname.Text = ""
creports.Text = ""
dactive
End Function

Private Sub add_Click()
add.Enabled = False
save.Enabled = True
flag = 1
active
End Sub
Private Sub delete_Click()
delete.Enabled = False
edit.Enabled = False
recordcheck
k = MsgBox("are you sure you want to delete", vbOKCancel)
If k = 1 Then
con.Execute ("delete  from creport where cno='" & sid & "'")
MsgBox ("Succesfully deleted")
clearit
fillgrid
dactive
recordcheck
Else
creport.Show
dactive
End If
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
rs.Open ("select * from creport where cno='" & sid & "'"), con, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
cname.Text = rs.Fields(1)
creports.Text = rs.Fields(2)
End If
dactive
add.Enabled = False
edit.Enabled = True
delete.Enabled = True
End Sub

Private Sub Form_Load()
connection
recordcheck
listing
fillgrid
dactive
add.Enabled = True
edit.Enabled = False
delete.Enabled = False
save.Enabled = False
clear.Enabled = True
End Sub

Private Sub ok_Click()
main.Show
Unload Me
End Sub

Private Sub save_Click()
If cname.Text = "" Or creports.Text = "" Then
 MsgBox ("Please enter all the fields")
Else
save.Enabled = False
recordcheck
If flag = 1 Then
con.Execute ("insert into creport values('" & UCase(cname.Text) & "','" & UCase(creports.Text) & "')")
MsgBox ("Succesfully saved")
fillgrid
clearit
dactive
End If
If flag = 2 Then
con.Execute ("update creport set cname='" & UCase(cname.Text) & "',creport='" & UCase(creports.Text) & "'")
MsgBox ("Successfuly edited")
fillgrid
recordcheck
clearit
dactive
End If
add.Enabled = True
edit.Enabled = False
delete.Enabled = False
End If
End Sub
