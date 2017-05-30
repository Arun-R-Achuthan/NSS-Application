VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form enroll 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Enrollment"
   ClientHeight    =   10500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17205
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   10500
   ScaleWidth      =   17205
   WindowState     =   2  'Maximized
   Begin VB.ComboBox eyear 
      Height          =   315
      ItemData        =   "Form1.frx":1323B
      Left            =   7920
      List            =   "Form1.frx":1325D
      TabIndex        =   47
      Top             =   7200
      Width           =   975
   End
   Begin MSComCtl2.DTPicker elast 
      Height          =   375
      Left            =   7920
      TabIndex        =   43
      Top             =   8400
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      Format          =   68943873
      CurrentDate     =   42271
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   600
      Top             =   960
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=nssfinal;Data Source=."
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=nssfinal;Data Source=."
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   2280
      Top             =   9720
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=nssfinal;Data Source=."
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=nssfinal;Data Source=."
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox eanswer 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   405
      Left            =   4680
      TabIndex        =   41
      Top             =   8400
      Width           =   1935
   End
   Begin VB.ComboBox equestion 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "Form1.frx":1329D
      Left            =   4680
      List            =   "Form1.frx":132AD
      TabIndex        =   39
      Top             =   7800
      Width           =   3495
   End
   Begin VB.TextBox id 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4680
      TabIndex        =   37
      Top             =   7200
      Width           =   2295
   End
   Begin MSFlexGridLib.MSFlexGrid flex 
      Height          =   6975
      Left            =   9720
      TabIndex        =   30
      Top             =   2160
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   12303
      _Version        =   393216
      Cols            =   5
      Appearance      =   0
   End
   Begin VB.ComboBox edistrict 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "Form1.frx":132F2
      Left            =   4680
      List            =   "Form1.frx":13320
      TabIndex        =   29
      Top             =   4080
      Width           =   2295
   End
   Begin VB.TextBox email 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      TabIndex        =   28
      Top             =   4560
      Width           =   3495
   End
   Begin VB.TextBox econtact 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      TabIndex        =   25
      Top             =   6720
      Width           =   2295
   End
   Begin VB.TextBox eheight 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      MaxLength       =   3
      TabIndex        =   23
      Top             =   6120
      Width           =   1335
   End
   Begin VB.TextBox eweight 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Enabled         =   0   'False
      Height          =   375
      Left            =   8040
      MaxLength       =   3
      TabIndex        =   21
      Top             =   6720
      Width           =   855
   End
   Begin VB.ComboBox ebgroup 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "Form1.frx":133B5
      Left            =   4680
      List            =   "Form1.frx":133D1
      TabIndex        =   19
      Top             =   5640
      Width           =   1335
   End
   Begin VB.ComboBox esem 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "Form1.frx":133F7
      Left            =   7560
      List            =   "Form1.frx":1340D
      TabIndex        =   16
      Top             =   5160
      Width           =   1335
   End
   Begin VB.ComboBox eclass 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "Form1.frx":13423
      Left            =   4680
      List            =   "Form1.frx":13436
      TabIndex        =   15
      Top             =   5160
      Width           =   1335
   End
   Begin VB.ComboBox egender 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "Form1.frx":13453
      Left            =   7560
      List            =   "Form1.frx":1345D
      TabIndex        =   13
      Top             =   5640
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker edob 
      Height          =   375
      Left            =   7560
      TabIndex        =   11
      Top             =   6120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   68943873
      CurrentDate     =   42271
   End
   Begin VB.TextBox epin 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      MaxLength       =   6
      TabIndex        =   9
      Top             =   3480
      Width           =   4215
   End
   Begin VB.TextBox elocation 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   2880
      Width           =   4215
   End
   Begin VB.TextBox ehouse 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   2280
      Width           =   4215
   End
   Begin VB.TextBox ename 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      MaxLength       =   20
      TabIndex        =   2
      Top             =   1680
      Width           =   4215
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "DOB"
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
      Left            =   6360
      TabIndex        =   46
      Top             =   6240
      Width           =   615
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Enroll yr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   7080
      TabIndex        =   45
      Top             =   7200
      Width           =   735
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   11640
      Top             =   9240
      Width           =   1095
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "  EXIT"
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
      Left            =   11760
      TabIndex        =   44
      Top             =   9360
      Width           =   1095
   End
   Begin VB.Label donate 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Donated"
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
      Left            =   6600
      TabIndex        =   42
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Answer"
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
      Left            =   3000
      TabIndex        =   40
      Top             =   8280
      Width           =   855
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Security Queston"
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
      Left            =   3000
      TabIndex        =   38
      Top             =   7680
      Width           =   1695
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
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
      Left            =   3000
      TabIndex        =   36
      Top             =   7200
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   10440
      Top             =   9240
      Width           =   1095
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   8880
      Top             =   9240
      Width           =   1455
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   7560
      Top             =   9240
      Width           =   975
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   495
      Left            =   6480
      Top             =   9240
      Width           =   975
   End
   Begin VB.Label save 
      BackStyle       =   0  'Transparent
      Caption         =   "   SAVE"
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
      Left            =   10440
      TabIndex        =   35
      Top             =   9360
      Width           =   1095
   End
   Begin VB.Label delete 
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
      ForeColor       =   &H80000009&
      Height          =   375
      Left            =   8880
      TabIndex        =   34
      Top             =   9360
      Width           =   1455
   End
   Begin VB.Label edit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "   EDIT"
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
      Left            =   7560
      TabIndex        =   33
      Top             =   9360
      Width           =   975
   End
   Begin VB.Label add 
      BackStyle       =   0  'Transparent
      Caption         =   "    ADD"
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
      Left            =   6480
      TabIndex        =   32
      Top             =   9360
      Width           =   975
   End
   Begin VB.Label Label17 
      Caption         =   "Label17"
      Height          =   15
      Left            =   4320
      TabIndex        =   31
      Top             =   9240
      Width           =   1695
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "District"
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
      Height          =   495
      Left            =   3000
      TabIndex        =   27
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact"
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
      Left            =   3000
      TabIndex        =   26
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "mail"
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
      Left            =   3000
      TabIndex        =   24
      Top             =   4560
      Width           =   495
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Weight"
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
      Left            =   7200
      TabIndex        =   22
      Top             =   6720
      Width           =   855
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Height"
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
      Left            =   3000
      TabIndex        =   20
      Top             =   6120
      Width           =   735
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Blood Group"
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
      Left            =   3000
      TabIndex        =   18
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Semester"
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
      Left            =   6360
      TabIndex        =   17
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Class"
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
      Left            =   3000
      TabIndex        =   14
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
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
      Left            =   6360
      TabIndex        =   12
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "DOB"
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
      Left            =   6360
      TabIndex        =   10
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "PIN"
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
      Left            =   3000
      TabIndex        =   8
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label5 
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
      Left            =   3000
      TabIndex        =   7
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "House Name"
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
      Left            =   3000
      TabIndex        =   4
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label3 
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
      Left            =   3000
      TabIndex        =   3
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Details"
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
      Height          =   15
      Left            =   2520
      TabIndex        =   1
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      Height          =   7455
      Left            =   2760
      Top             =   1560
      Width           =   6255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enrollment"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   8760
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "enroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag As Integer
Dim i As Integer
Dim j As Integer
Dim sid As Integer
Public Function active()
ename.Enabled = True
ehouse.Enabled = True
elocation.Enabled = True
epin.Enabled = True
edistrict.Enabled = True
email.Enabled = True
eclass.Enabled = True
esem.Enabled = True
ebgroup.Enabled = True
egender.Enabled = True
eheight.Enabled = True
edob.Enabled = True
econtact.Enabled = True
eweight.Enabled = True
id.Enabled = True
equestion.Enabled = True
eanswer.Enabled = True
eyear.Enabled = True
End Function
Public Function dactive()
ename.Enabled = False
ehouse.Enabled = False
elocation.Enabled = False
epin.Enabled = False
edistrict.Enabled = False
email.Enabled = False
eclass.Enabled = False
esem.Enabled = False
ebgroup.Enabled = False
egender.Enabled = False
eheight.Enabled = False
edob.Enabled = False
econtact.Enabled = False
eweight.Enabled = False
id.Enabled = False
equestion.Enabled = False
eanswer.Enabled = False
eyear.Enabled = False
End Function

Public Function clearit()
ename.Text = ""
ehouse.Text = ""
elocation.Text = ""
epin.Text = ""
edistrict.Text = ""
email.Text = ""
eclass.Text = ""
esem.Text = ""
ebgroup.Text = ""
egender.Text = ""
eheight.Text = ""
econtact.Text = ""
eweight.Text = ""
id.Text = ""
equestion.Text = ""
eanswer.Text = ""
eyear.Text = ""
End Function
Public Function fillgrid()
flex.TextMatrix(0, 0) = "No"
flex.TextMatrix(0, 1) = "Name"
flex.TextMatrix(0, 2) = "Bgroup"
flex.TextMatrix(0, 3) = "Contact"
flex.TextMatrix(0, 4) = "ID"
recordcheck

rs.Open ("select eno,ename,ebgroup,econtact,id from enroll"), con, adOpenDynamic, adLockOptimistic
j = 1
flex.Rows = 1

While (Not rs.EOF)
flex.Rows = flex.Rows + 1
flex.TextMatrix(j, 0) = Trim(rs.Fields(0))
flex.TextMatrix(j, 1) = Trim(rs.Fields(1))
flex.TextMatrix(j, 2) = Trim(rs.Fields(2))
flex.TextMatrix(j, 3) = Trim(rs.Fields(3))
flex.TextMatrix(j, 4) = Trim(rs.Fields(4))
j = j + 1
rs.MoveNext
Wend
End Function


Private Sub add_Click()
active
flag = 1
add.Enabled = False
save.Enabled = True
End Sub

Private Sub clear_Click()
clearit
dactive
End Sub

Private Sub delete_Click()
recordcheck
If ident = 2 Then
MsgBox ("You dont have the access")
Else
k = MsgBox("Are you sure you want to delete", vbOKCancel)
If k = 1 Then
con.Execute ("delete from enroll where eno='" & sid & "'")
MsgBox ("successfully deleted")
dactive
fillgrid
clearit
Else
enroll.Show
dactive
clearit
End If
End If
delete.Enabled = False
edit.Enabled = False
End Sub
Private Sub eanswer_LostFocus()
If Not ValidName(eanswer.Text) Then
 MsgBox ("Please enter a valid answer")
 eanswer.SetFocus
End If
End Sub

Private Sub econtact_LostFocus()
If Not ValidPhone(econtact.Text) Then
    MsgBox "Please enter a correct phone number"
    econtact.SetFocus
End If

    
End Sub

Private Sub edit_Click()
flag = 2
active
edit.Enabled = False
save.Enabled = True
delete.Enabled = False
End Sub







Private Sub eheight_LostFocus()
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
Else
KeyAscii = 0
End If
End Sub

Private Sub ehouse_LostFocus()
If Not ValidName(ehouse.Text) Then
   MsgBox " Please enter a correct house name "
   elocation.SetFocus
   End If
End Sub

Private Sub elocation_LostFocus()
If Not ValidName(elocation.Text) Then
   MsgBox " Please enter a correct location "
   elocation.SetFocus
   End If
End Sub





Private Sub email_LostFocus()
If Not ValidEmail(email.Text) Then
  MsgBox "Please enter a valid email"
  email.SetFocus
 End If
End Sub

Private Sub ename_LostFocus()
If Not ValidName(ename.Text) Then
   MsgBox " Please enter a correct name "
   ename.SetFocus
   End If
End Sub
Private Sub epin_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
Else
KeyAscii = 0
End If

End Sub

Private Sub epin_LostFocus()
If Len(epin.Text) < 6 Then
MsgBox ("Enter a valid pin")
epin.SetFocus
End If
End Sub
Private Sub eweight_LostFocus()
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
Else
KeyAscii = 0
End If
End Sub



Private Sub eyear_LostFocus()
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
Else
KeyAscii = 0
End If
End Sub

Private Sub flex_Click()
sid = flex.TextMatrix(flex.Row, 0)
recordcheck
rs.Open ("select * from enroll where eno='" & sid & "'")
ename.Text = rs.Fields(1)
ehouse.Text = rs.Fields(2)
elocation.Text = rs.Fields(3)
epin.Text = rs.Fields(4)
edistrict.Text = rs.Fields(5)
email.Text = rs.Fields(6)
eclass.Text = rs.Fields(7)
esem.Text = rs.Fields(8)
ebgroup.Text = rs.Fields(9)
egender.Text = rs.Fields(10)
eheight.Text = rs.Fields(11)
edob.Value = rs.Fields(12)
econtact.Text = rs.Fields(13)
eweight.Text = rs.Fields(14)
id.Text = rs.Fields(15)
equestion.Text = rs.Fields(16)
eanswer.Text = rs.Fields(17)
eyear.Text = rs.Fields(18)
elast.Value = rs.Fields(21)
dactive
add.Enabled = False
edit.Enabled = True
delete.Enabled = True
End Sub

Private Sub Form_Load()
connection
fillgrid
dactive
add.Enabled = True
save.Enabled = False
delete.Enabled = False
edit.Enabled = False

End Sub

Private Sub Label18_Click()
Unload Me
main.Show
End Sub

Private Sub save_Click()
If ename.Text = "" Or epin.Text = "" Or ehouse.Text = "" Or elocation.Text = "" Or epin.Text = "" Or econtact.Text = "" Or eheight.Text = "" Or email.Text = "" Or egender.Text = "" Or eclass.Text = "" Or esem.Text = "" Or email.Text = "" Or equestion.Text = "" Or eanswer.Text = "" Then
MsgBox ("Please enter all the fields")
Else
recordcheck
If flag = 1 Then
con.Execute ("insert into enroll values ('" & UCase(ename.Text) & "','" & UCase(ehouse.Text) & "','" & UCase(elocation.Text) & "','" & UCase(epin.Text) & "','" & UCase(edistrict.Text) & "','" & LCase(email.Text) & "','" & UCase(eclass.Text) & "','" & UCase(esem.Text) & "','" & UCase(ebgroup.Text) & "','" & UCase(egender.Text) & "','" & UCase(eheight.Text) & "','" & edob.Value & "','" & UCase(econtact.Text) & "','" & UCase(eweight.Text) & "','" & UCase(id.Text) & "','" & UCase(equestion.Text) & "','" & LCase(eanswer.Text) & "','" & UCase(eyear.Text) & "','" & UCase(id.Text) & "','" & UCase(id.Text) & "','" & elast.Value & "')")
MsgBox ("Successfully saved")
fillgrid
clearit
End If
If flag = 2 Then
con.Execute ("update enroll set ename='" & UCase(ename.Text) & "',ehouse='" & UCase(ehouse.Text) & "',elocation='" & UCase(elocation.Text) & "',epin='" & UCase(epin.Text) & "',edistrict='" & UCase(edistrict.Text) & "',email='" & LCase(email.Text) & "',esem= '" & UCase(esem.Text) & "',ebgroup='" & ebgroup.Text & "', egender='" & UCase(egender.Text) & "', eheight='" & eheight.Text & "',edob='" & edob.Value & "', econtact= '" & econtact.Text & "',eweight='" & eweight.Text & "',equestion='" & UCase(equestion.Text) & "',eanswer='" & LCase(eanswer.Text) & "',eyear='" & eyear.Text & "',elast='" & elast.Value & "' where sid='" & eno & "'")
fillgrid
MsgBox ("Successfully Edited")
clearit
End If
End If
save.Enabled = False
add.Enabled = True
edit.Enabled = False
End Sub




