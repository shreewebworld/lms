VERSION 5.00
Begin VB.Form frmdeletemember 
   BackColor       =   &H00000000&
   Caption         =   "Delete Member"
   ClientHeight    =   5250
   ClientLeft      =   7605
   ClientTop       =   3210
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   6795
   Begin VB.TextBox txtmphone 
      BackColor       =   &H00FFFFFF&
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
      Left            =   2640
      TabIndex        =   14
      Top             =   3600
      Width           =   2535
   End
   Begin VB.TextBox txtmaddr 
      BackColor       =   &H00FFFFFF&
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
      Left            =   2640
      TabIndex        =   12
      Top             =   3000
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton btnnext 
      BackColor       =   &H00FFC0FF&
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton btnprev 
      BackColor       =   &H00FFC0FF&
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton btncancel 
      BackColor       =   &H0080FF80&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton btndelete 
      BackColor       =   &H008080FF&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox txtmemail 
      BackColor       =   &H00FFFFFF&
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
      Left            =   2640
      TabIndex        =   6
      Top             =   2400
      Width           =   2535
   End
   Begin VB.TextBox txtmname 
      BackColor       =   &H00FFFFFF&
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
      Left            =   2640
      TabIndex        =   5
      Top             =   1800
      Width           =   2535
   End
   Begin VB.TextBox txtmemberid 
      BackColor       =   &H00FFFFFF&
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
      Left            =   2640
      TabIndex        =   4
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label P 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   1440
      TabIndex        =   15
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   1320
      TabIndex        =   13
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   600
      Picture         =   "frmdeletemember.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   960
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Member ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Name "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Update or Delete Member"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   5175
   End
End
Attribute VB_Name = "frmdeletemember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim strquery As String
Private Sub btndelete_Click()
msg = MsgBox("Delete Member :" & rs!mname, vbYesNo, "Delete_Confirm")
If msg = vbYes Then
rs.Delete
MsgBox ("Member Record Deleted Successfully")
End If
End Sub

Private Sub btncancel_Click()
Unload Me
End Sub

Private Sub btnnext_Click()
rs.MoveNext
If rs.EOF Then
rs.MoveFirst
End If
txtmemberid.Text = rs.Fields("memberid")
txtmname.Text = rs.Fields("mname")
txtmemail.Text = rs.Fields("memail")
txtmaddr.Text = rs.Fields("maddr")
txtmphone.Text = rs.Fields("mphone")
End Sub

Private Sub btnprev_Click()
rs.MovePrevious
If rs.BOF Then
rs.MoveLast
End If
txtmemberid.Text = rs.Fields("memberid")
txtmname.Text = rs.Fields("mname")
txtmemail.Text = rs.Fields("memail")
txtmaddr.Text = rs.Fields("maddr")
txtmphone.Text = rs.Fields("mphone")
End Sub

Private Sub Command1_Click()

rs!mname = txtmname.Text
rs!memail = txtmemail.Text
rs!maddr = txtmaddr.Text
rs!mphone = txtmphone.Text
rs.Update
a = MsgBox("Updated Member Information", vbInformation)
End Sub

Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data source=" & App.Path & "\clgdata.mdb;Persist Security Info=false"
strquery = "Select*from members"
rs.Open strquery, con, adOpenDynamic, adLockOptimistic
txtmemberid.Text = rs.Fields("memberid")
txtmname.Text = rs.Fields("mname")
txtmemail.Text = rs.Fields("memail")
txtmaddr.Text = rs.Fields("maddr")
txtmphone.Text = rs.Fields("mphone")
End Sub

Private Sub Form_Unload(Cancel As Integer)
rs.Close
con.Close
End Sub
