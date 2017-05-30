VERSION 5.00
Begin VB.Form board 
   Caption         =   "Board"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "board.frx":0000
   ScaleHeight     =   10500
   ScaleWidth      =   18165
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   360
      Top             =   2880
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000009&
      Height          =   2415
      Left            =   4560
      Top             =   240
      Width           =   8535
   End
   Begin VB.Image Image1 
      Height          =   2175
      Left            =   4680
      Picture         =   "board.frx":1323B
      Stretch         =   -1  'True
      Top             =   360
      Width           =   8295
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      Height          =   615
      Left            =   7800
      Shape           =   4  'Rounded Rectangle
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label log 
      BackStyle       =   0  'Transparent
      Caption         =   "        Login"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   7800
      TabIndex        =   10
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Image k4 
      Height          =   3975
      Left            =   5640
      Picture         =   "board.frx":29EC2
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   6375
   End
   Begin VB.Image k3 
      Height          =   3975
      Left            =   5640
      Picture         =   "board.frx":3C391
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   6375
   End
   Begin VB.Image k2 
      Height          =   3975
      Left            =   5640
      Picture         =   "board.frx":57133
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   6375
   End
   Begin VB.Image k1 
      Height          =   3975
      Left            =   5640
      Picture         =   "board.frx":B1130
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   6375
   End
   Begin VB.Label l4 
      BackStyle       =   0  'Transparent
      Caption         =   "BPC College floor 3 room number 56"
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
      Height          =   615
      Left            =   12240
      TabIndex        =   9
      Top             =   5760
      Width           =   3735
   End
   Begin VB.Label l66 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact 9786865744"
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
      Left            =   1680
      TabIndex        =   8
      Top             =   6120
      Width           =   2655
   End
   Begin VB.Label l55 
      BackStyle       =   0  'Transparent
      Caption         =   "Shijin George (Joint secretary)"
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
      Left            =   1680
      TabIndex        =   7
      Top             =   5640
      Width           =   3375
   End
   Begin VB.Label l44 
      BackStyle       =   0  'Transparent
      Caption         =   "Amalumol Ealias (Secretary)"
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
      Left            =   1680
      TabIndex        =   6
      Top             =   5040
      Width           =   3255
   End
   Begin VB.Label l33 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact : 9890998789"
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
      Left            =   1680
      TabIndex        =   5
      Top             =   4560
      Width           =   3015
   End
   Begin VB.Label l22 
      BackStyle       =   0  'Transparent
      Caption         =   "Prof Sheba KU (Prog. Officer)"
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
      Left            =   1680
      TabIndex        =   4
      Top             =   3960
      Width           =   3735
   End
   Begin VB.Label l11 
      BackStyle       =   0  'Transparent
      Caption         =   "Prof MJ Kurian (Prog. Officer)"
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
      Left            =   1680
      TabIndex        =   3
      Top             =   3360
      Width           =   3495
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   3975
      Left            =   1320
      Top             =   2880
      Width           =   4215
   End
   Begin VB.Label l3 
      BackStyle       =   0  'Transparent
      Caption         =   "www.facebook.com/nss_bpc.com"
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
      Left            =   12240
      TabIndex        =   2
      Top             =   5040
      Width           =   3855
   End
   Begin VB.Label l2 
      BackStyle       =   0  'Transparent
      Caption         =   "www.nss_bpc.com"
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
      Left            =   12240
      TabIndex        =   1
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      Height          =   3975
      Left            =   12120
      Top             =   2880
      Width           =   4215
   End
   Begin VB.Label l1 
      BackStyle       =   0  'Transparent
      Caption         =   "nss.nic.in"
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
      Left            =   12240
      TabIndex        =   0
      Top             =   3360
      Width           =   1215
   End
End
Attribute VB_Name = "board"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Private Sub Form_Load()
i = 0
End Sub

Private Sub log_Click()
login.Show
Unload Me
End Sub

Private Sub Timer1_Timer()
i = i + 1
If l11.Top <= 3000 Then
l11.Top = 6120
l11.Left = 1680
End If
If l22.Top <= 3000 Then
l22.Top = 6120
l22.Left = 1680
End If
If l33.Top <= 3000 Then
l33.Top = 6120
l33.Left = 1680
End If
If l44.Top <= 3000 Then
l44.Top = 6120
l44.Left = 1680
End If
If l55.Top <= 3000 Then
l55.Top = 6120
l55.Left = 1680
End If
If l66.Top <= 3000 Then
l66.Top = 6120
l66.Left = 1680
End If
If l4.Top <= 3000 Then
l4.Top = 6200
l4.Left = 12240
End If
If l3.Top <= 3000 Then
l3.Top = 6200
l3.Left = 12240
End If
If l2.Top <= 3000 Then
l2.Top = 6200
l2.Left = 12240
End If
If l1.Top <= 3000 Then
l1.Top = 6200
l1.Left = 12240
End If
l1.Top = l1.Top - 200
l2.Top = l2.Top - 200
l3.Top = l3.Top - 200
l4.Top = l4.Top - 200
l11.Top = l11.Top - 200
l22.Top = l22.Top - 200
l33.Top = l33.Top - 200
l44.Top = l44.Top - 200
l55.Top = l55.Top - 200
l66.Top = l66.Top - 200
If i = 1 Then
k1.Visible = True
k2.Visible = False
k3.Visible = False
k4.Visible = False
End If
If i = 2 Then
k1.Visible = False
k2.Visible = True
k3.Visible = False
k4.Visible = False
End If
If i = 3 Then
k3.Visible = True
k1.Visible = False
k2.Visible = False
k4.Visible = False
End If
If i = 4 Then
k4.Visible = True
k1.Visible = False
k2.Visible = False
k3.Visible = False
i = 0
End If

End Sub
