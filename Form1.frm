VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "LMS"
   ClientHeight    =   8415
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17370
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form1.frx":119A2
   ScaleHeight     =   8415
   ScaleWidth      =   17370
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4560
      Top             =   6840
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3720
      Top             =   6840
   End
   Begin VB.CommandButton btnaddbook 
      BackColor       =   &H0080C0FF&
      Caption         =   "+ Add Book"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   8040
      Width           =   17370
      _ExtentX        =   30639
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "10:27 AM"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "10/18/2015"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7057
            MinWidth        =   7057
            Text            =   "Project Developed by Dnyaneshwar Naiknavare"
            TextSave        =   "Project Developed by Dnyaneshwar Naiknavare"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton btnaddmember 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "+ Add Member"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton btnreturn 
      BackColor       =   &H0080C0FF&
      Caption         =   "Return Book"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton btnissue 
      BackColor       =   &H0080C0FF&
      Caption         =   "Issue Book"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3120
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   6240
      TabIndex        =   1
      Top             =   3120
      Width           =   7095
      Begin VB.CommandButton Command2 
         BackColor       =   &H000080FF&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CommandButton btnlogin 
         BackColor       =   &H0000FF00&
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox txtpassword 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox txtuserid 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   2
         Top             =   840
         Width           =   2295
      End
      Begin VB.Image Image1 
         Height          =   1920
         Left            =   4800
         Picture         =   "Form1.frx":F9F30
         Top             =   600
         Width           =   1920
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   6
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "User ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   5
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Library Managament System"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1365
      Left            =   1830
      TabIndex        =   0
      Top             =   360
      Width           =   17190
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim strquery As String

Private Sub btnaddbook_Click()
frmaddbook.Show
End Sub

Private Sub btnlogin_Click()
Dim flag As Boolean
flag = False

con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data source=" & App.Path & "\clgdata.mdb;Persist Security Info=False"
strquery = "Select*from users"
rs.Open strquery, con, adOpenDynamic, adLockOptimistic

rs.MoveFirst
While Not rs.EOF
If txtuserid.Text = rs!userid And txtpassword.Text = rs!Password Then
flag = True
user = rs!uname
End If
rs.MoveNext
Wend
If flag = True Then
a = MsgBox("Welcome " & user, vbOKOnly, "Login Success")
txtuserid.Text = ""
txtpassword.Text = ""
MDIForm1.mnulogin1.Enabled = False
MDIForm1.Enabled = True
MDIForm1.mnubooks.Enabled = True
MDIForm1.mnumembers.Enabled = True
MDIForm1.mnureport.Enabled = True
MDIForm1.mnuhelp.Enabled = True
MDIForm1.mnutransaction.Enabled = True
MDIForm1.mnusearch.Enabled = True
btnissue.Visible = True
btnreturn.Visible = True
btnaddmember.Visible = True
btnaddbook.Visible = True
Frame1.Visible = False
Timer1.Enabled = True
Else
MsgBox ("Invalid UserID or Password ")
End If

rs.Close
con.Close
End Sub

Private Sub Command2_Click()
Frame1.Visible = False
End Sub

Private Sub btnissue_Click()
frmbookissue.Show
End Sub

Private Sub btnreturn_Click()
frmreturnbook.Show
End Sub

Private Sub btnaddmember_Click()
frmaddmember.Show
End Sub


Private Sub Timer1_Timer()
If btnaddbook.Width < 6000 Then
btnreturn.Width = btnreturn.Width + 200
btnissue.Width = btnissue.Width + 200
btnaddbook.Width = btnaddbook.Width + 200
btnaddmember.Width = btnaddmember.Width + 200
Else
Timer1.Enabled = False
Timer2.Enabled = True
End If
End Sub

Private Sub Timer2_Timer()
If btnissue.Width > 3495 Then
btnreturn.Width = btnreturn.Width - 150
btnissue.Width = btnissue.Width - 150
btnaddbook.Width = btnaddbook.Width - 150
btnaddmember.Width = btnaddmember.Width - 150
Else
Timer2.Enabled = False
End If
End Sub


Private Sub txtpassword_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
Dim flag As Boolean
flag = False

con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data source=" & App.Path & "\clgdata.mdb;Persist Security Info=False"
strquery = "Select*from users"
rs.Open strquery, con, adOpenDynamic, adLockOptimistic

rs.MoveFirst
While Not rs.EOF
If txtuserid.Text = rs!userid And txtpassword.Text = rs!Password Then
flag = True
user = rs!uname
End If
rs.MoveNext
Wend
If flag = True Then
a = MsgBox("Welcome " & user, vbOKOnly, "Login Success")
txtuserid.Text = ""
txtpassword.Text = ""
MDIForm1.mnulogin1.Enabled = False
MDIForm1.Enabled = True
MDIForm1.mnubooks.Enabled = True
MDIForm1.mnumembers.Enabled = True
MDIForm1.mnureport.Enabled = True
MDIForm1.mnuhelp.Enabled = True
MDIForm1.mnutransaction.Enabled = True
MDIForm1.mnusearch.Enabled = True
btnissue.Visible = True
btnreturn.Visible = True
btnaddmember.Visible = True
btnaddbook.Visible = True
Frame1.Visible = False
Timer1.Enabled = True
Else
MsgBox ("Invalid UserID or Password ")
End If

rs.Close
con.Close
End If
End Sub
