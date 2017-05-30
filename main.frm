VERSION 5.00
Begin VB.MDIForm main 
   BackColor       =   &H8000000C&
   Caption         =   "NSS management"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   Picture         =   "main.frx":0000
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   4680
      TabIndex        =   1
      Top             =   15
      Width           =   4680
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   4680
      TabIndex        =   0
      Top             =   0
      Width           =   4680
   End
   Begin VB.Menu mnuadmin 
      Caption         =   "Admin"
      Begin VB.Menu mnuactivity 
         Caption         =   "Activity"
         Begin VB.Menu mnuadd 
            Caption         =   "Add"
         End
         Begin VB.Menu mnuattend 
            Caption         =   "Attendence"
         End
         Begin VB.Menu mnusreport 
            Caption         =   "Report"
         End
      End
      Begin VB.Menu mnucamp 
         Caption         =   "Camp"
         Begin VB.Menu mnuaddcamp 
            Caption         =   "Add"
         End
         Begin VB.Menu mnucampreport 
            Caption         =   "Report"
         End
      End
      Begin VB.Menu mnublood 
         Caption         =   "Blood "
         Begin VB.Menu mnubsearch 
            Caption         =   "Blood Search"
         End
         Begin VB.Menu mnubqueue 
            Caption         =   "Blood Requests"
         End
         Begin VB.Menu blist 
            Caption         =   "Blood Donation Entry"
         End
      End
      Begin VB.Menu mnuenroll 
         Caption         =   "Enrollment"
      End
      Begin VB.Menu mnufundrequest 
         Caption         =   "Fund Request"
      End
      Begin VB.Menu mnureportss 
         Caption         =   "Reports"
      End
      Begin VB.Menu mnuvsearch 
         Caption         =   "Volunteer Search"
      End
   End
   Begin VB.Menu mnuvol 
      Caption         =   "Volunteer"
      Begin VB.Menu mnuvolactivity 
         Caption         =   "Activity"
      End
      Begin VB.Menu mnuvolcamp 
         Caption         =   "Camp"
      End
      Begin VB.Menu mnuchpass 
         Caption         =   "Change Password"
      End
      Begin VB.Menu mnuvolfreq 
         Caption         =   "Financial Request"
      End
      Begin VB.Menu mnuvolbreq 
         Caption         =   "Blood Request"
      End
   End
   Begin VB.Menu mnuout 
      Caption         =   "Sign out"
   End
   Begin VB.Menu mnuexit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub blist_Click()
bloodlist.Show
End Sub

Private Sub MDIForm_Load()
If nos = 1 Then
mnuvol.Enabled = False
End If
If nos = 2 Then
mnuadmin.Enabled = False
End If
End Sub

Private Sub mnuadd_Click()
activity.Show
End Sub

Private Sub mnuaddcamp_Click()
camp.Show
End Sub

Private Sub mnuattend_Click()
activityattend.Show
End Sub

Private Sub mnubqueue_Click()
brequest.Show
End Sub

Private Sub mnubsearch_Click()
bsearch.Show
End Sub

Private Sub mnucampreport_Click()
creport.Show
End Sub

Private Sub mnuchpass_Click()
chpassword.Show
End Sub

Private Sub mnuenroll_Click()
enroll.Show
End Sub

Private Sub mnuexit_Click()
End
End Sub
Private Sub mnufundrequest_Click()
fundraise.Show
End Sub

Private Sub mnuout_Click()
k = MsgBox("Are you sure you want to Sign out", vbOKCancel)
If k = 1 Then
MsgBox ("Signed out")
load.Show
Unload Me
Else
main.Show
End If
End Sub

Private Sub mnureport_Click()
areports.Show
End Sub

Private Sub mnureportss_Click()
reports.Show
End Sub

Private Sub mnusreport_Click()
areports.Show
End Sub

Private Sub mnuvolactivity_Click()
activity.Show
End Sub

Private Sub mnuvolbreq_Click()
brequest.Show
End Sub

Private Sub mnuvolcamp_Click()
camp.Show
End Sub
Private Sub mnuvolfreq_Click()
fundraise.Show
End Sub

Private Sub mnuvsearch_Click()
vsearch.Show
End Sub
