VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form areports 
   Caption         =   "Activity Report"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   Picture         =   "areport.frx":0000
   ScaleHeight     =   10500
   ScaleWidth      =   18165
   WindowState     =   2  'Maximized
   Begin VB.TextBox aname 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   6120
      TabIndex        =   13
      Top             =   2400
      Width           =   2415
   End
   Begin VB.ComboBox adates 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6120
      TabIndex        =   11
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox areport 
      Appearance      =   0  'Flat
      Height          =   4335
      Left            =   6120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   3000
      Width           =   3735
   End
   Begin MSFlexGridLib.MSFlexGrid flex 
      Height          =   6495
      Left            =   10200
      TabIndex        =   1
      Top             =   1080
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   11456
      _Version        =   393216
      Rows            =   1
      Appearance      =   0
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Activity Name"
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
      Left            =   4080
      TabIndex        =   12
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Activity Date"
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
      Left            =   4080
      TabIndex        =   10
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   13080
      Top             =   7800
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
      Left            =   13080
      TabIndex        =   9
      Top             =   7920
      Width           =   1095
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   12000
      Top             =   7800
      Width           =   975
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   10680
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   9240
      Top             =   7800
      Width           =   1335
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   8040
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   6840
      Top             =   7800
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
      Left            =   12000
      TabIndex        =   8
      Top             =   7920
      Width           =   975
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
      Left            =   9240
      TabIndex        =   7
      Top             =   7920
      Width           =   1335
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
      Left            =   10680
      TabIndex        =   6
      Top             =   7920
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
      Left            =   8040
      TabIndex        =   5
      Top             =   7920
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
      Left            =   6840
      TabIndex        =   4
      Top             =   7920
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Activity Report"
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
      Left            =   4080
      TabIndex        =   2
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Activity Report"
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
      Left            =   9120
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      Height          =   6495
      Left            =   3960
      Top             =   1080
      Width           =   6015
   End
End
Attribute VB_Name = "areports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim flag As Integer
Dim sid As Integer
Public Function listing()
recordcheck
rs.Open ("select * from activityin "), con, adOpenDynamic, adLockOptimistic
While rs.EOF = False
adates.AddItem (rs.Fields(2))
rs.MoveNext
Wend
End Function
Public Function fillgrid()
flex.TextMatrix(0, 0) = "No"
flex.TextMatrix(0, 1) = "Date"
recordcheck
rs.Open ("select ano,adetails,adate from areport"), con, adOpenDynamic, adLockOptimistic
flex.Rows = 1
i = 1
While (rs.EOF = False)
flex.Rows = flex.Rows + 1
flex.TextMatrix(i, 0) = rs.Fields(0)
flex.TextMatrix(i, 1) = rs.Fields(2)
rs.MoveNext
i = i + 1
Wend
End Function
Public Function clearit()
areport.Text = ""
adates.Text = ""
End Function
Public Function active()
areport.Enabled = True
adates.Enabled = True
End Function
Public Function dactive()
areport.Enabled = False
adates.Enabled = False

End Function




Private Sub adates_LostFocus()
If adates.Text <> "" Then
 recordcheck
 rs.Open ("select aname from activityin where adate = '" & adates.Text & "'"), con, adOpenDynamic, adLockOptimistic
 aname.Text = rs.Fields(0)
 End If
End Sub



Private Sub aname_KeyPress(KeyAscii As Integer)
If KeyAscii > 0 Then
 KeyAscii = 0
End If
End Sub

Private Sub flex_Click()
sid = flex.TextMatrix(flex.Row, 0)
recordcheck
rs.Open ("select * from areport where ano='" & sid & "'"), con, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
adates.Text = rs.Fields(2)
areport.Text = rs.Fields(1)
recordcheck
rs.Open ("select aname from activityin where adate = '" & adates.Text & "'"), con, adOpenDynamic, adLockOptimistic
aname.Text = rs.Fields(0)
End If
dactive
add.Enabled = False
delete.Enabled = True
edit.Enabled = True
clear.Enabled = False
End Sub

Private Sub Form_Load()
connection
recordcheck
fillgrid
dactive
listing
add.Enabled = True
edit.Enabled = False
delete.Enabled = False
save.Enabled = False
clear.Enabled = True
End Sub

Private Sub add_Click()
add.Enabled = False
save.Enabled = True
flag = 1
active
End Sub

Private Sub edit_Click()
edit.Enabled = False
delete.Enabled = False
save.Enabled = True
clear.Enabled = False
flag = 2
active
End Sub

Private Sub clear_Click()
clearit
dactive
End Sub

Private Sub delete_Click()
recordcheck
k = MsgBox("are you sure you want to delete", vbOKCancel)
If k = 1 Then
con.Execute ("delete  from areport where ano='" & sid & "'")
MsgBox ("Succesfully deleted")
clearit
fillgrid
dactive
recordcheck
Else
areports.Show
dactive
End If
delete.Enabled = False
add.Enabled = True
edit.Enabled = False
End Sub



Private Sub save_Click()
If areport.Text = "" Or adates.Text = "" Then
 MsgBox ("Please enter all the fields")
Else
recordcheck
If flag = 1 Then
con.Execute ("insert into areport values('" & UCase(areport.Text) & "','" & adates.Text & "')")
MsgBox ("Succesfully saved")
fillgrid
clearit
dactive
End If
If flag = 2 Then
con.Execute ("update areport set adetails='" & UCase(areport.Text) & "' adate='" & adates.Text & "'")
MsgBox ("Successfuly edited")
fillgrid
recordcheck
clearit
dactive
End If
save.Enabled = False
add.Enabled = True
clear.Enabled = True
End If
End Sub



Private Sub ok_Click()
main.Show
Unload Me
End Sub
