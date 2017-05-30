VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vsearch 
   Caption         =   "Volunteer Search"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "vsearch.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   2715
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flex 
      Height          =   3735
      Left            =   10200
      TabIndex        =   4
      Top             =   3720
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   6588
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      Appearance      =   0
   End
   Begin VB.ComboBox vid 
      Height          =   315
      Left            =   7200
      TabIndex        =   1
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      Height          =   615
      Left            =   7440
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label ok 
      BackStyle       =   0  'Transparent
      Caption         =   "     EXIT"
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
      Left            =   7440
      TabIndex        =   5
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label KASDHJSJK 
      BackStyle       =   0  'Transparent
      Caption         =   "Volunteer Search"
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
      Left            =   8280
      TabIndex        =   3
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   615
      Left            =   5520
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label search 
      BackStyle       =   0  'Transparent
      Caption         =   "   SEARCH"
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
      Left            =   5520
      TabIndex        =   2
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Voulunteer ID"
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
      Left            =   4800
      TabIndex        =   0
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000004&
      Height          =   3735
      Left            =   4560
      Top             =   3720
      Width           =   5415
   End
End
Attribute VB_Name = "vsearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag As Integer
Dim i As Integer
Dim sid As Integer
Dim f As Integer
Public Function listing()
recordcheck
rs.Open ("select * from enroll "), con, adOpenDynamic, adLockOptimistic
While rs.EOF = False
vid.AddItem (rs.Fields(15))
rs.MoveNext
Wend
End Function
Public Function fillgrid()
flex.TextMatrix(0, 0) = "ID"
flex.TextMatrix(0, 1) = "Name"
flex.TextMatrix(0, 2) = "Class"
flex.TextMatrix(0, 3) = "Sem"
flex.TextMatrix(0, 4) = "Contact"
recordcheck
rs.Open ("select id,ename,eclass,esem,econtact from enroll where id = '" & vid.Text & "'"), con, adOpenDynamic, adLockOptimistic
flex.Rows = 1
i = 1
While (rs.EOF = False)
f = 2
flex.Rows = flex.Rows + 1
flex.TextMatrix(i, 0) = rs.Fields(0)
flex.TextMatrix(i, 1) = rs.Fields(1)
flex.TextMatrix(i, 2) = rs.Fields(2)
flex.TextMatrix(i, 3) = rs.Fields(3)
flex.TextMatrix(i, 4) = rs.Fields(4)
rs.MoveNext
i = i + 1
Wend
If f = 2 Then
MsgBox ("Found")
Else
MsgBox ("Not Found")
End If
End Function



Private Sub Form_Load()
connection
recordcheck
f = 0
listing
End Sub

Private Sub ok_Click()
Unload Me
main.Show
End Sub

Private Sub search_Click()
fillgrid
End Sub
