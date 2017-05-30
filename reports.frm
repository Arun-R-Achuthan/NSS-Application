VERSION 5.00
Begin VB.Form reports 
   Caption         =   "Report"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "reports.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.ComboBox b6 
      Height          =   315
      Left            =   14400
      TabIndex        =   21
      Top             =   6840
      Width           =   1575
   End
   Begin VB.ComboBox bgroup 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "reports.frx":34FFD
      Left            =   14160
      List            =   "reports.frx":34FFF
      TabIndex        =   18
      Top             =   1560
      Width           =   1335
   End
   Begin VB.ComboBox vid 
      Height          =   315
      Left            =   8640
      TabIndex        =   14
      Top             =   6840
      Width           =   1575
   End
   Begin VB.ComboBox cnames 
      Height          =   315
      Left            =   3000
      TabIndex        =   11
      Top             =   6960
      Width           =   1935
   End
   Begin VB.ComboBox rdates 
      Height          =   315
      Left            =   7920
      TabIndex        =   5
      Top             =   1560
      Width           =   2535
   End
   Begin VB.ComboBox radate 
      Height          =   315
      Left            =   2880
      TabIndex        =   1
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Shape Shape14 
      BorderColor     =   &H80000007&
      Height          =   615
      Left            =   8280
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Shape Shape13 
      BorderColor     =   &H80000007&
      Height          =   615
      Left            =   6120
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "        EXIT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   495
      Left            =   8280
      TabIndex        =   25
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "     Next Page"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   495
      Left            =   6120
      TabIndex        =   24
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Shape Shape12 
      BorderColor     =   &H80000006&
      Height          =   615
      Left            =   13200
      Top             =   7920
      Width           =   2175
   End
   Begin VB.Label printnotice 
      BackStyle       =   0  'Transparent
      Caption         =   "     Print Notice"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   495
      Left            =   13200
      TabIndex        =   23
      Top             =   8040
      Width           =   2175
   End
   Begin VB.Label Da 
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
      ForeColor       =   &H80000007&
      Height          =   375
      Left            =   12120
      TabIndex        =   22
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Activity Notice"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   375
      Left            =   13320
      TabIndex        =   20
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Shape Shape11 
      BorderColor     =   &H80000006&
      Height          =   3615
      Left            =   11640
      Shape           =   4  'Rounded Rectangle
      Top             =   5880
      Width           =   5175
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H80000006&
      Height          =   615
      Left            =   13320
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label printlist 
      BackStyle       =   0  'Transparent
      Caption         =   "     Print List"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   495
      Left            =   13320
      TabIndex        =   19
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label5 
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
      ForeColor       =   &H80000007&
      Height          =   375
      Left            =   12000
      TabIndex        =   17
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H80000006&
      Height          =   615
      Left            =   7680
      Top             =   7920
      Width           =   2415
   End
   Begin VB.Label printdetail 
      BackStyle       =   0  'Transparent
      Caption         =   "        Print Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   495
      Left            =   7680
      TabIndex        =   16
      Top             =   8040
      Width           =   2415
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Class"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   495
      Left            =   6480
      TabIndex        =   15
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Volunteer Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   495
      Left            =   7560
      TabIndex        =   13
      Top             =   6120
      Width           =   2055
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H80000006&
      Height          =   3495
      Left            =   6240
      Shape           =   4  'Rounded Rectangle
      Top             =   5880
      Width           =   4815
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H80000006&
      Height          =   615
      Left            =   2160
      Top             =   7920
      Width           =   2295
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "  Generate Report"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   375
      Left            =   2160
      TabIndex        =   12
      Top             =   8040
      Width           =   2295
   End
   Begin VB.Label Label8 
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
      ForeColor       =   &H80000007&
      Height          =   375
      Left            =   1440
      TabIndex        =   10
      Top             =   6960
      Width           =   855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Camp Report"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H80000006&
      Height          =   3495
      Left            =   960
      Shape           =   4  'Rounded Rectangle
      Top             =   5880
      Width           =   4815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Blood List"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   375
      Left            =   13080
      TabIndex        =   8
      Top             =   720
      Width           =   1815
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000006&
      Height          =   3615
      Left            =   11760
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   4815
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000006&
      Height          =   615
      Left            =   7440
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Label rgreport2 
      BackStyle       =   0  'Transparent
      Caption         =   "      Print Report"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   495
      Left            =   7440
      TabIndex        =   7
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label Label4 
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
      ForeColor       =   &H80000007&
      Height          =   375
      Left            =   6600
      TabIndex        =   6
      Top             =   1560
      Width           =   855
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
      ForeColor       =   &H80000007&
      Height          =   375
      Left            =   7800
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000006&
      Height          =   3495
      Left            =   6360
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   4815
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000006&
      Height          =   615
      Left            =   2040
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label rgreport 
      BackStyle       =   0  'Transparent
      Caption         =   "   Print Attendence"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H80000007&
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Activity Attendence"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   720
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000006&
      Height          =   3495
      Left            =   840
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   4815
   End
End
Attribute VB_Name = "reports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Private Sub Form_Load()
connection
recordcheck
rs.Open ("select * from activityin "), con, adOpenDynamic, adLockOptimistic
While rs.EOF = False
radate.AddItem (rs.Fields(2))
rs.MoveNext
Wend
recordcheck

rs.Open ("select * from activityin "), con, adOpenDynamic, adLockOptimistic
While rs.EOF = False
rdates.AddItem (rs.Fields(2))
rs.MoveNext
Wend
recordcheck
rs.Open ("select  distinct eclass from enroll "), con, adOpenDynamic, adLockOptimistic
While rs.EOF = False
vid.AddItem (rs.Fields(0))
rs.MoveNext
Wend
recordcheck

rs.Open (" select distinct(ebgroup) from enroll "), con, adOpenDynamic, adLockOptimistic
While rs.EOF = False
bgroup.AddItem (rs.Fields(0))
rs.MoveNext
Wend
bgroup.Text = "[ Select Blood group ]"
recordcheck
rs.Open ("select * from camp"), con, adOpenDynamic, adLockOptimistic
While rs.EOF = False
cnames.AddItem (rs.Fields(1))
rs.MoveNext
Wend
recordcheck
rs.Open ("select * from activityin "), con, adOpenDynamic, adLockOptimistic
While rs.EOF = False
b6.AddItem (rs.Fields(2))
rs.MoveNext
Wend
End Sub




Private Sub Label13_Click()
Unload Me
reports1.Show
End Sub

Private Sub Label14_Click()
Unload Me

End Sub

Private Sub Label9_Click()
If DataEnvironment1.Connection1.State = 1 Then
DataEnvironment1.Connection1.Close
End If
DataEnvironment1.Connection1.Open
campreport.Show
DataEnvironment1.Command5 vid.Text
End Sub

Private Sub printdetail_Click()
If DataEnvironment1.Connection1.State = 1 Then
DataEnvironment1.Connection1.Close
End If
DataEnvironment1.Connection1.Open
volreport.Show
DataEnvironment1.Command3 vid.Text
End Sub

Private Sub printlist_Click()
If DataEnvironment1.Connection1.State = 1 Then
DataEnvironment1.Connection1.Close
End If
DataEnvironment1.Connection1.Open
blist.Show
DataEnvironment1.Command4 bgroup.Text
End Sub

Private Sub printnotice_Click()
If DataEnvironment1.Connection1.State = 1 Then
DataEnvironment1.Connection1.Close
End If
DataEnvironment1.Connection1.Open
prints.Show
DataEnvironment1.Command6 b6.Text
End Sub

Private Sub rgreport_Click()
If DataEnvironment1.Connection1.State = 1 Then
DataEnvironment1.Connection1.Close
End If
DataEnvironment1.Connection1.Open
attendence.Show
DataEnvironment1.Command1 radate.Text
End Sub
Private Sub rgreport2_Click()
If DataEnvironment1.Connection1.State = 1 Then
DataEnvironment1.Connection1.Close
End If
DataEnvironment1.Connection1.Open
areportss.Show
DataEnvironment1.Command2 rdates.Text
End Sub
