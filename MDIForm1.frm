VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Library Management System"
   ClientHeight    =   5175
   ClientLeft      =   4890
   ClientTop       =   3480
   ClientWidth     =   11265
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Menu mnulogin 
      Caption         =   "Login"
      Begin VB.Menu mnulogin1 
         Caption         =   "Login"
      End
      Begin VB.Menu mnulogout 
         Caption         =   "Log out"
      End
      Begin VB.Menu hypen 
         Caption         =   "-"
      End
      Begin VB.Menu mnunewuser 
         Caption         =   "New User"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnutransaction 
      Caption         =   "Transaction"
      Begin VB.Menu mnuissuebook 
         Caption         =   "Issue Book"
      End
      Begin VB.Menu mnureturnbook 
         Caption         =   "Return Book"
      End
   End
   Begin VB.Menu mnusearch 
      Caption         =   "Search"
      Begin VB.Menu mnusearchbook 
         Caption         =   "Book"
      End
      Begin VB.Menu mnusearchmember 
         Caption         =   "Member"
      End
   End
   Begin VB.Menu mnubooks 
      Caption         =   "Books"
      Begin VB.Menu mnuadd 
         Caption         =   "Add Book"
      End
      Begin VB.Menu mnudelete 
         Caption         =   "Delete Book"
      End
      Begin VB.Menu mnuupdate 
         Caption         =   "Update Book"
      End
   End
   Begin VB.Menu mnumembers 
      Caption         =   "Members"
      Begin VB.Menu mnuaddmember 
         Caption         =   "Add Member"
      End
      Begin VB.Menu mnuupdatemember 
         Caption         =   "Update Member"
      End
      Begin VB.Menu mnudeletemember 
         Caption         =   "Delete Member"
      End
   End
   Begin VB.Menu mnureport 
      Caption         =   "Report"
      Begin VB.Menu mnubook 
         Caption         =   "Books"
      End
      Begin VB.Menu mnumember 
         Caption         =   "Members"
      End
      Begin VB.Menu mnuissuedbooks 
         Caption         =   "Issued Books"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "Help"
      Begin VB.Menu mnuabout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_Load()
MDIForm1.mnubooks.Enabled = False
MDIForm1.mnumembers.Enabled = False
MDIForm1.mnureport.Enabled = False
MDIForm1.mnuhelp.Enabled = False
MDIForm1.mnutransaction.Enabled = False
MDIForm1.mnusearch.Enabled = False
frmmain.Show
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
End
End Sub

Private Sub mnuabout_Click()
frmabout.Show
End Sub

Private Sub mnuadd_Click()
frmaddbook.Show
End Sub

Private Sub mnuaddmember_Click()
frmaddmember.Show
End Sub

Private Sub mnubook_Click()
frmbookreport.Show
End Sub

Private Sub mnudelete_Click()
frmdeletebook.Show
End Sub

Private Sub mnudeletemember_Click()
frmdeletemember.Show
End Sub

Private Sub mnuexit_Click()
End
End Sub

Private Sub mnuissuebook_Click()
frmbookissue.Show
End Sub

Private Sub mnuissuedbooks_Click()
frmissuedbooksreport.Show
End Sub

Private Sub mnulogin1_Click()
frmmain.Show
frmmain.Frame1.Visible = True
End Sub

Private Sub mnulogout_Click()
MDIForm1.mnutransaction.Visible = False
MDIForm1.mnubooks.Enabled = False
MDIForm1.mnumembers.Enabled = False
MDIForm1.mnureport.Enabled = False
MDIForm1.mnuhelp.Enabled = False
MDIForm1.mnusearch.Enabled = False
frmmain.btnissue.Visible = False
frmmain.btnreturn.Visible = False
frmmain.btnaddbook.Visible = False
frmmain.btnaddmember.Visible = False
frmmain.Frame1.Visible = True
mnulogin1.Enabled = True
End Sub

Private Sub mnumember_Click()
frmmembersreport.Show
End Sub

Private Sub mnunewuser_Click()
frmnewuser.Show
End Sub

Private Sub mnureturnbook_Click()
frmreturnbook.Show
End Sub


Private Sub mnusearchbook_Click()
frmsearchb.Show
End Sub

Private Sub mnusearchmember_Click()
frmsearchm.Show
End Sub

Private Sub mnuupdate_Click()
frmupdatebook.Show
End Sub

Private Sub mnuupdatemember_Click()
frmdeletemember.Show
End Sub
