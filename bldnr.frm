VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form bldnr 
   Caption         =   "Blood Donar"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "bldnr.frx":0000
   ScaleHeight     =   10500
   ScaleWidth      =   18165
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox dnames 
      Height          =   285
      Left            =   5640
      TabIndex        =   17
      Top             =   5040
      Width           =   2655
   End
   Begin MSFlexGridLib.MSFlexGrid flex 
      Height          =   5055
      Left            =   9480
      TabIndex        =   10
      Top             =   2640
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   8916
      _Version        =   393216
      Rows            =   1
      Cols            =   6
      Appearance      =   0
   End
   Begin VB.ComboBox hname 
      Height          =   315
      Left            =   5640
      TabIndex        =   9
      Top             =   6240
      Width           =   1695
   End
   Begin VB.ComboBox dod 
      Height          =   315
      Left            =   5640
      TabIndex        =   7
      Top             =   5640
      Width           =   1695
   End
   Begin VB.ComboBox bgrp 
      Height          =   315
      Left            =   5640
      TabIndex        =   6
      Top             =   4440
      Width           =   1695
   End
   Begin VB.ComboBox rname 
      Height          =   315
      Left            =   5640
      TabIndex        =   5
      Top             =   3840
      Width           =   2655
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   12600
      Top             =   8160
      Width           =   975
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H80000009&
      BorderColor     =   &H80000005&
      Height          =   495
      Left            =   11520
      Top             =   8160
      Width           =   975
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   10200
      Top             =   8160
      Width           =   1215
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   8880
      Top             =   8160
      Width           =   1215
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   7560
      Top             =   8160
      Width           =   1215
   End
   Begin VB.Shape add 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   6120
      Top             =   8160
      Width           =   1335
   End
   Begin VB.Label ok 
      BackStyle       =   0  'Transparent
      Caption         =   "    OK"
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
      Left            =   12600
      TabIndex        =   16
      Top             =   8280
      Width           =   975
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
      Left            =   11520
      TabIndex        =   15
      Top             =   8280
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
      Left            =   10200
      TabIndex        =   14
      Top             =   8280
      Width           =   1215
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
      Left            =   8880
      TabIndex        =   13
      Top             =   8280
      Width           =   1215
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
      Top             =   8280
      Width           =   1215
   End
   Begin VB.Label adds 
      BackStyle       =   0  'Transparent
      Caption         =   "     ADD"
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
      Top             =   8280
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Hospital Name"
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
      Left            =   3600
      TabIndex        =   8
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Blood Donation"
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
      Left            =   6840
      TabIndex        =   4
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of donation"
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
      Left            =   3600
      TabIndex        =   3
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Donar Name"
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
      Left            =   3600
      TabIndex        =   2
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label Label2 
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
      Left            =   3600
      TabIndex        =   1
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Receiver Name"
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
      Left            =   3600
      TabIndex        =   0
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000004&
      Height          =   5055
      Left            =   3240
      Top             =   2640
      Width           =   6015
   End
End
Attribute VB_Name = "bldnr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim falg As Integer
Dim i As Integer
Dim sid As Integer
Public Function listing()

End Function
Public Function clearit()
rname.Text = ""
dnames.Text = ""
dod.Text = ""
hname.Text = ""
bgrp.Text = ""
End Function
Public Function active()
rname.Enabled = True
dnames.Enabled = True
dod.Enabled = True
hname.Enabled = True
bgrp.Enabled = True
End Function
Public Function dactive()
rname.Enabled = False
dnames.Enabled = False
dod.Enabled = False
hname.Enabled = False
bgrp.Enabled = False
End Function
Public Function fillgrid()
flex.TextMatrix(0, 0) = "No"
flex.TextMatrix(0, 1) = "Donar Name"
flex.TextMatrix(0, 2) = "Receiver Name"
flex.TextMatrix(0, 3) = "Date"
flex.TextMatrix(0, 4) = "Hospital Name"
flex.TextMatrix(0, 5) = "Blood Group"
recordcheck
rs.Open ("select blno,blrname,bldname,bldate,blgrp,hname from bldonate"), con, adOpenDynamic, adLockOptimistic
flex.Rows = 1
i = 1
While (rs.EOF = False)
flex.Rows = flex.Rows + 1
flex.TextMatrix(i, 0) = rs.Fields(0)
flex.TextMatrix(i, 1) = rs.Fields(2)
flex.TextMatrix(i, 2) = rs.Fields(1)
flex.TextMatrix(i, 3) = rs.Fields(3)
flex.TextMatrix(i, 4) = rs.Fields(4)
flex.TextMatrix(i, 5) = rs.Fields(5)
rs.MoveNext
i = i + 1
Wend
End Function


Private Sub adds_Click()
flag = 1
active
End Sub

Private Sub clear_Click()
clearit
dactive
End Sub

Private Sub edit_Click()
flag = 2
active
End Sub

Private Sub Form_Load()
connection
recordcheck
fillgrid
listing
End Sub

Private Sub Label8_Click()

End Sub

Private Sub ok_Click()
Unload Me
main.Show
End Sub

Private Sub save_Click()
recordcheck
If flag = 1 Then
con.Execute ("insert into bldonate values('" & rname.Text & "','" & dnames.Text & "','" & dod.Text & "','" & bgrp.Text & "','" & hname.Text & "')")
MsgBox ("Succesfully saved")
fillgrid
clearit
dactive
End If
If flag = 2 Then
con.Execute ("update bldonate set blrname ='" & rname.Text & "',bldname='" & dnames.Text & "',bldate='" & dod.Text & "',blgrp='" & bgrp.Text & "',hname='" & hname.Text & "' where ano ='" & sid & "'")
MsgBox ("Successfuly edited")
fillgrid
recordcheck
clearit
dactive
End If
End Sub
