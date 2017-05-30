VERSION 5.00
Begin VB.Form activityattend 
   Caption         =   "Activity Attendence"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "activityattend.frx":0000
   MDIChild        =   -1  'True
   Picture         =   "activityattend.frx":000C
   ScaleHeight     =   10500
   ScaleWidth      =   18165
   WindowState     =   2  'Maximized
   Begin VB.ComboBox id 
      Height          =   315
      Left            =   8160
      TabIndex        =   8
      Top             =   4800
      Width           =   2055
   End
   Begin VB.ComboBox adate 
      Height          =   315
      Left            =   8160
      TabIndex        =   7
      Top             =   5520
      Width           =   2055
   End
   Begin VB.ComboBox aname 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8160
      TabIndex        =   2
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   9000
      Top             =   6480
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
      Height          =   375
      Left            =   9000
      TabIndex        =   6
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label djaskd 
      BackStyle       =   0  'Transparent
      Caption         =   "Activity Attendence"
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
      Left            =   7320
      TabIndex        =   5
      Top             =   2760
      Width           =   3015
   End
   Begin VB.Label D 
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
      Height          =   495
      Left            =   6120
      TabIndex        =   4
      Top             =   5520
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   7080
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Label submit 
      BackStyle       =   0  'Transparent
      Caption         =   "    SUBMIT"
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
      TabIndex        =   3
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Label Label2 
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
      Left            =   6120
      TabIndex        =   1
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label lab 
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
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      Height          =   4095
      Left            =   5640
      Top             =   3480
      Width           =   6015
   End
End
Attribute VB_Name = "activityattend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim k As Integer
Private Sub Form_Load()
k = 0
connection
recordcheck
rs.Open ("select * from activityin "), con, adOpenDynamic, adLockOptimistic
While rs.EOF = False
aname.AddItem (rs.Fields(1))
adate.AddItem (rs.Fields(2))
rs.MoveNext
Wend
recordcheck
rs.Open ("select id from enroll"), con, adOpenDynamic, adLockOptimistic
While rs.EOF = False
id.AddItem (rs.Fields(0))
rs.MoveNext
Wend
End Sub

Private Sub Label1_Click()
If DataEnvironment1.Connection2.State = 1 Then
DataEnvironment1.Connection2.Close
End If
DataEnvironment1.Connection2.Open
DataReport2.Show
DataEnvironment1.comand2 adate.Text
End Sub



Private Sub submit_Click()
If aname.Text = "" Or adate.Text = "" Or id.Text = "" Then
 MsgBox ("Please enter all fields")
Else
recordcheck
rs.Open ("select id from activityattend where adate = '" & adate.Text & "'"), con, adOpenDynamic, adLockOptimistic
While Not rs.EOF
 If id.Text = rs.Fields(0) Then
  k = 1
 End If
rs.MoveNext
Wend
If k = 0 Then
con.Execute ("insert into activityattend values('" & aname.Text & "','" & id.Text & "','" & adate.Text & "')")
MsgBox ("Succesfully saved")
adate.Text = ""
aname.Text = ""
id.Text = ""
Else
MsgBox (" Attendence already marked for this id")
End If
End If
End Sub

Private Sub ok_Click()
main.Show
Unload Me
End Sub

