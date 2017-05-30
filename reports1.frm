VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form reports1 
   BackColor       =   &H00808000&
   Caption         =   "Reports1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "reports1.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.ComboBox bnames 
      Height          =   315
      Left            =   14520
      TabIndex        =   10
      Top             =   2160
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker d2 
      Height          =   495
      Left            =   8640
      TabIndex        =   7
      Top             =   2400
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      _Version        =   393216
      Format          =   68943873
      CurrentDate     =   42283
   End
   Begin MSComCtl2.DTPicker d1 
      Height          =   495
      Left            =   8640
      TabIndex        =   6
      Top             =   1680
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      _Version        =   393216
      CalendarBackColor=   -2147483635
      Format          =   68943873
      CurrentDate     =   42283
   End
   Begin VB.ComboBox fnames 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2280
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H80000006&
      Height          =   615
      Left            =   7560
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label ok 
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
      ForeColor       =   &H80000006&
      Height          =   495
      Left            =   7560
      TabIndex        =   13
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H80000006&
      Height          =   615
      Left            =   13320
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "      Print Request"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   495
      Left            =   13320
      TabIndex        =   12
      Top             =   3360
      Width           =   2415
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
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   12480
      TabIndex        =   11
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label7 
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
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   13440
      TabIndex        =   9
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000007&
      Height          =   3735
      Left            =   12000
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   4815
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000007&
      Height          =   615
      Left            =   7800
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label6 
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
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   7800
      TabIndex        =   8
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Date between"
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
      TabIndex        =   5
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Blood Donation List"
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
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000007&
      Height          =   3615
      Left            =   6360
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   4815
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000007&
      Height          =   615
      Left            =   1560
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "     Print Request"
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
      Left            =   1680
      TabIndex        =   3
      Top             =   3360
      Width           =   2295
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
      ForeColor       =   &H80000007&
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Financial Request "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000007&
      Height          =   3615
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   4815
   End
End
Attribute VB_Name = "reports1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
connection
recordcheck
rs.Open ("select fname from fund"), con, adOpenDynamic, adLockOptimistic
While rs.EOF = False
fnames.AddItem (rs.Fields(0))
rs.MoveNext
Wend
recordcheck
rs.Open ("select bname from blood"), con, adOpenDynamic, adLockOptimistic
While rs.EOF = False
bnames.AddItem (rs.Fields(0))
rs.MoveNext
Wend
End Sub

Private Sub Label3_Click()
If DataEnvironment1.Connection1.State = 1 Then
DataEnvironment1.Connection1.Close
End If
DataEnvironment1.Connection1.Open
fundreport.Show
DataEnvironment1.Command7 fnames.Text
End Sub

Private Sub Label6_Click()
If DataEnvironment1.Connection1.State = 1 Then
DataEnvironment1.Connection1.Close
End If
DataEnvironment1.Connection1.Open
listreport.Show
DataEnvironment1.Command8 d1.Value, d2.Value
End Sub

Private Sub Label9_Click()
If DataEnvironment1.Connection1.State = 1 Then
DataEnvironment1.Connection1.Close
End If
DataEnvironment1.Connection1.Open
brequestrep.Show
DataEnvironment1.Command9 bnames.Text
End Sub

Private Sub ok_Click()
Unload Me
End Sub
