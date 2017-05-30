VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form bsearch 
   Caption         =   "Blood Search"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "bsearch.frx":0000
   ScaleHeight     =   10500
   ScaleWidth      =   18165
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid flex 
      Height          =   3255
      Left            =   8280
      TabIndex        =   7
      Top             =   1200
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   5741
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      Appearance      =   0
   End
   Begin VB.ComboBox bsdistrict 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "bsearch.frx":1323B
      Left            =   4920
      List            =   "bsearch.frx":13269
      TabIndex        =   6
      Top             =   3720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ComboBox bsgroup 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "bsearch.frx":132FE
      Left            =   4920
      List            =   "bsearch.frx":1331A
      TabIndex        =   4
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   615
      Left            =   8880
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "      OK"
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
      Left            =   8880
      TabIndex        =   9
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000009&
      Height          =   615
      Left            =   7320
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "   Search"
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
      TabIndex        =   8
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label cc 
      BackStyle       =   0  'Transparent
      Caption         =   "District"
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
      TabIndex        =   5
      Top             =   3720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label b 
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
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Shape s2 
      BorderColor     =   &H80000010&
      Height          =   615
      Left            =   5040
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Shape s1 
      BorderColor     =   &H80000010&
      Height          =   615
      Left            =   2880
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label cussearch 
      BackStyle       =   0  'Transparent
      Caption         =   "   Custom Search"
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
      Left            =   5040
      TabIndex        =   2
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label expsearch 
      BackStyle       =   0  'Transparent
      Caption         =   " Express Search"
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
      Left            =   2880
      TabIndex        =   1
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      Height          =   3255
      Left            =   2760
      Top             =   1200
      Width           =   5055
   End
   Begin VB.Label Blo 
      BackStyle       =   0  'Transparent
      Caption         =   "Blood Search"
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
      Left            =   7080
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "bsearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i, j As Integer
Dim flag As Integer
Dim found As Integer
Dim y As Variant
Dim c As Variant



Private Sub cussearch_Click()
b.Visible = True
bsgroup.Visible = True
bsdistrict.Visible = True
cc.Visible = True
flag = 2
End Sub

Private Sub expsearch_Click()
b.Visible = True
bsgroup.Visible = True
cc.Visible = False
bsdistrict.Visible = False
flag = 1
End Sub

Private Sub Form_Load()
flag = 0
found = 0
j = 1
bsgroup.Visible = False
bsdistrict.Visible = False
cc.Visible = False
b.Visible = False
connection
End Sub
Public Function fillgrid()
j = 1
connection
recordcheck
flex.Rows = 1
If flag = 1 Then
rs.Open ("select * from enroll where ebgroup = '" & bsgroup.Text & "' and  DATEDIFF(D,elast,'" & Date & "') > 60"), con, adOpenDynamic, adLockOptimistic
found = 1
While Not rs.EOF
flex.TextMatrix(0, 0) = "Name"
flex.TextMatrix(0, 1) = "Location"
flex.TextMatrix(0, 2) = "District"
flex.TextMatrix(0, 3) = "Contact"
flex.TextMatrix(0, 4) = "Gender"

flex.Rows = flex.Rows + 1
flex.TextMatrix(j, 0) = rs.Fields(0)
flex.TextMatrix(j, 1) = rs.Fields(1)
flex.TextMatrix(j, 2) = rs.Fields(2)
flex.TextMatrix(j, 3) = rs.Fields(3)
flex.TextMatrix(j, 4) = rs.Fields(4)
rs.MoveNext
j = j + 1
flag = -1
Wend
End If
recordcheck
If flag = 2 Then
rs.Open ("select * from enroll where ebgroup = '" & bsgroup.Text & "' and edistrict = '" & bsdistrict.Text & "'   and  DATEDIFF(D,elast,'" & Date & "') > 60"), con, adOpenDynamic, adLockOptimistic
found = 1
While Not rs.EOF
flex.TextMatrix(0, 0) = "Name"
flex.TextMatrix(0, 1) = "Location"
flex.TextMatrix(0, 2) = "District"
flex.TextMatrix(0, 3) = "Contact"
flex.TextMatrix(0, 4) = "Gender"

flex.Rows = flex.Rows + 1
flex.TextMatrix(j, 0) = rs.Fields(0)
flex.TextMatrix(j, 1) = rs.Fields(1)
flex.TextMatrix(j, 2) = rs.Fields(2)
flex.TextMatrix(j, 3) = rs.Fields(3)
flex.TextMatrix(j, 4) = rs.Fields(4)
rs.MoveNext
j = j + 1
flag = -1
Wend
End If
If found = 1 Then
MsgBox ("Found")
Else
MsgBox ("Not found")
End If
End Function



Private Sub Label1_Click()
main.Show
Unload Me
End Sub

Private Sub Label5_Click()
fillgrid
found = 0
End Sub
