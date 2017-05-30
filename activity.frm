VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form activity 
   Caption         =   "Activity"
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9675
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "activity.frx":0000
   ScaleHeight     =   4650
   ScaleWidth      =   9675
   WindowState     =   2  'Maximized
   Begin VB.TextBox adetails 
      Appearance      =   0  'Flat
      Height          =   1575
      Left            =   4800
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   4800
      Width           =   3855
   End
   Begin VB.TextBox alocation 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   4800
      TabIndex        =   8
      Top             =   4080
      Width           =   3855
   End
   Begin VB.TextBox atime 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   4800
      MaxLength       =   7
      TabIndex        =   7
      Top             =   3360
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker adate 
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      Top             =   2640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   16580609
      CurrentDate     =   42271
   End
   Begin VB.TextBox aname 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   1920
      Width           =   3855
   End
   Begin MSFlexGridLib.MSFlexGrid flex 
      Height          =   6255
      Left            =   9480
      TabIndex        =   0
      Top             =   1320
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   11033
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      Appearance      =   0
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H80000009&
      Height          =   615
      Left            =   8520
      Top             =   8760
      Width           =   1455
   End
   Begin VB.Label Ex 
      BackStyle       =   0  'Transparent
      Caption         =   "      EXIT"
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
      Left            =   8520
      TabIndex        =   17
      Top             =   8880
      Width           =   1455
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H80000009&
      Height          =   615
      Left            =   10920
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   615
      Left            =   9840
      Top             =   7920
      Width           =   975
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000009&
      Height          =   615
      Left            =   8520
      Top             =   7920
      Width           =   975
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      Height          =   615
      Left            =   7320
      Top             =   7920
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   615
      Left            =   6120
      Top             =   7920
      Width           =   1095
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
      Height          =   495
      Left            =   10920
      TabIndex        =   16
      Top             =   8040
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
      Height          =   735
      Left            =   9840
      TabIndex        =   15
      Top             =   7800
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
      Height          =   735
      Left            =   8520
      TabIndex        =   14
      Top             =   7800
      Width           =   975
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
      Height          =   495
      Left            =   7320
      TabIndex        =   13
      Top             =   8040
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
      Height          =   495
      Left            =   6120
      TabIndex        =   12
      Top             =   8040
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Activity"
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
      Left            =   8640
      TabIndex        =   11
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3840
      TabIndex        =   10
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3840
      TabIndex        =   2
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3840
      TabIndex        =   1
      Top             =   1920
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      Height          =   6255
      Left            =   3600
      Top             =   1320
      Width           =   5415
   End
End
Attribute VB_Name = "activity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim sid As Integer
Dim flag As Integer
Dim k As Integer
Public Function active()
adate.Enabled = True
alocation.Enabled = True
atime.Enabled = True
aname.Enabled = True
adetails.Enabled = True
End Function
Public Function dactive()
adate.Enabled = False
alocation.Enabled = False
atime.Enabled = False
aname.Enabled = False
adetails.Enabled = False
End Function

Public Function fillgrid()
flex.TextMatrix(0, 0) = "No"
flex.TextMatrix(0, 1) = "Name"
flex.TextMatrix(0, 2) = "Date"
flex.TextMatrix(0, 3) = "Location"
flex.TextMatrix(0, 4) = "Time"
recordcheck
rs.Open ("select ano,aname,adate,alocation,atime from activityin"), con, adOpenDynamic, adLockOptimistic
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
aname.Text = ""
alocation.Text = ""
atime.Text = ""
adetails.Text = ""
End Function

Private Sub add_Click()
flag = 1
active
save.Enabled = True
add.Enabled = False
End Sub





Private Sub alocation_LostFocus()
If Not ValidName(alocation.Text) Then
 MsgBox ("Enter a valid location")
 alocation.SetFocus
End If
End Sub

Private Sub aname_LostFocus()
If Not ValidName(aname.Text) Then
 MsgBox ("Please enter a correct name")
 aname.SetFocus
 End If
End Sub



Private Sub atime_Click()
MsgBox ("Enter in the fromat 9.00 AM")
End Sub

Private Sub clear_Click()
clearit
dactive
End Sub

Private Sub delete_Click()
recordcheck
k = MsgBox("are you sure you want to delete", vbOKCancel)
If k = 1 Then
con.Execute ("delete  from bldonate where bno='" & sid & "'")
MsgBox ("Succesfully deleted")
clearit
fillgrid
dactive
recordcheck
Else
activity.Show
dactive
End If
delete.Enabled = False
add.Enabled = True
End Sub

Private Sub edit_Click()
flag = 2
active
edit.Enabled = False
save.Enabled = True
delete.Enabled = False
End Sub

Private Sub Exit_Click()
main.Show
Unload Me
End Sub

Private Sub Ex_Click()
main.Show
Unload Me
End Sub

Private Sub flex_Click()

sid = flex.TextMatrix(flex.Row, 0)
recordcheck
rs.Open ("select * from activityin where ano='" & sid & "'"), con, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
aname.Text = rs.Fields(1)
adate.Value = rs.Fields(2)
alocation.Text = rs.Fields(3)
atime.Text = rs.Fields(4)
adetails.Text = rs.Fields(5)
dactive
If nos = 1 Then
delete.Enabled = True
edit.Enabled = True
add.Enabled = False
clear.Enabled = False
save.Enabled = False
End If
End If

End Sub

Private Sub Form_Load()
k = 0
If nos = 2 Then
add.Enabled = False
edit.Enabled = False
delete.Enabled = False
clear.Enabled = False
save.Enabled = False
End If
If nos = 1 Then
add.Enabled = True
edit.Enabled = False
delete.Enabled = False
clear.Enabled = False
save.Enabled = False
Ex.Enabled = True
End If
connection
rs.Open ("select * from activityin "), con, adOpenDynamic, adLockOptimistic
i = 1
recordcheck
fillgrid
dactive

End Sub



Private Sub save_Click()
If aname.Text = "" Or adate.Value = "" Or adetails.Text = "" Or atime.Text = "" Or alocation.Text = "" Then
 MsgBox ("Enter all fields")
Else
recordcheck
If flag = 1 Then
con.Execute ("insert into activityin values('" & UCase(aname.Text) & "','" & adate.Value & "','" & UCase(alocation.Text) & "','" & UCase(atime.Text) & "','" & UCase(adetails.Text) & "')")
MsgBox ("Succesfully saved")
fillgrid
clearit
dactive
End If
If flag = 2 Then
con.Execute ("update activityin set aname='" & UCase(aname.Text) & "',adate='" & adate.Value & "',alocation='" & UCase(alocation.Text) & "',atime='" & UCase(atime.Text) & "',adetails='" & UCase(adetails.Text) & "'where ano = '" & sid & "'")
MsgBox ("Successfuly edited")
fillgrid
recordcheck
clearit
dactive
End If
save.Enabled = False
add.Enabled = True
End If
End Sub
