VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form duty 
   Caption         =   "Duty Leave"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "duty.frx":0000
   ScaleHeight     =   10500
   ScaleWidth      =   18165
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid flex 
      Height          =   4335
      Left            =   9240
      TabIndex        =   9
      Top             =   2160
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   7646
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.TextBox reasons 
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   5160
      TabIndex        =   8
      Top             =   5520
      Width           =   3615
   End
   Begin MSComCtl2.DTPicker dates 
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   4560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   69271553
      CurrentDate     =   42287
   End
   Begin VB.TextBox names 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      Top             =   2760
      Width           =   3495
   End
   Begin VB.ComboBox ids 
      Height          =   315
      Left            =   5160
      TabIndex        =   5
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   11640
      Top             =   7080
      Width           =   975
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   10560
      Top             =   7080
      Width           =   975
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   9360
      Top             =   7080
      Width           =   1095
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   7800
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   6720
      Top             =   7080
      Width           =   975
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   5520
      Top             =   7080
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "   OK"
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
      Left            =   11640
      TabIndex        =   15
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label Label9 
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
      Left            =   10560
      TabIndex        =   14
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label Label8 
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
      Left            =   9360
      TabIndex        =   13
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "  DELETE"
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
      Left            =   7800
      TabIndex        =   12
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "  EDIT"
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
      Left            =   6720
      TabIndex        =   11
      Top             =   7200
      Width           =   975
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
      Left            =   5520
      TabIndex        =   10
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Reason"
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
      Left            =   3960
      TabIndex        =   4
      Top             =   5400
      Width           =   975
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
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
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
      Left            =   3960
      TabIndex        =   2
      Top             =   3600
      Width           =   495
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
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   2760
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      Height          =   4335
      Left            =   3720
      Top             =   2160
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Duty Leave"
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
      Left            =   8040
      TabIndex        =   0
      Top             =   1080
      Width           =   1815
   End
End
Attribute VB_Name = "duty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
dactive
connection
recordcheck
End Sub
Public Function clearit()
names.Text = ""
ids.Text = ""
reasons.Text = ""
End Function

Public Function active()
names.Enabled = True
dates.Enabled = True
ids.Enabled = True
reasons.Enabled = True
End Function
Public Function dactive()
names.Enabled = False
dates.Enabled = False
ids.Enabled = False
reasons.Enabled = False
End Function



