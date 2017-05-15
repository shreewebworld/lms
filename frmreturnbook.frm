VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmreturnbook 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Return Book"
   ClientHeight    =   7200
   ClientLeft      =   7170
   ClientTop       =   2640
   ClientWidth     =   5820
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   5820
   Begin VB.TextBox txtfine 
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
      TabIndex        =   17
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   240
      Top             =   0
   End
   Begin VB.CommandButton btncancel 
      BackColor       =   &H0080C0FF&
      Caption         =   "Cancel"
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
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton btnreturn 
      BackColor       =   &H0000FF00&
      Caption         =   "Return"
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
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6600
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   2640
      TabIndex        =   12
      Top             =   5160
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   52428801
      CurrentDate     =   42222
   End
   Begin VB.TextBox txtmemberid 
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
      TabIndex        =   11
      Top             =   3960
      Width           =   2055
   End
   Begin VB.TextBox txtmname 
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
      TabIndex        =   10
      Top             =   3360
      Width           =   2055
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   4680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   52428801
      CurrentDate     =   42222
   End
   Begin VB.TextBox txtbname 
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
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Enter Book Id and Press Enter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1335
      Left            =   480
      TabIndex        =   0
      Top             =   1200
      Width           =   4815
      Begin VB.TextBox txtbid 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   2
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "BookID"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   495
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Fine"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   16
      Top             =   5760
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   840
      Picture         =   "frmreturnbook.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   900
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Return Book"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   495
      Left            =   1920
      TabIndex        =   15
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "M_ID"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Member Name"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Return Date"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   7
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Issue Date"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Book Name"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
End
Attribute VB_Name = "frmreturnbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim strquery As String

Private Sub btnreturn_Click()
If txtbid.Text = "" Then
MsgBox ("Enter Book ID & press Enter")
Else

con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data source=" & App.Path & "\clgdata.mdb;Persist Security Info=false"
strquery = "Select*from transactionn"
rs.Open strquery, con, adOpenDynamic, adLockOptimistic
rs.MoveFirst
While Not rs.EOF
If rs!bid = Val(txtbid.Text) Then
rs!bookstatus = "Returned"
rs!return_date = DTPicker2.Value
rs!fine = txtfine.Text
rs.Update
flag = True
End If
rs.MoveNext
Wend
rs.Close
con.Close

con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data source=" & App.Path & "\clgdata.mdb;Persist Security Info=false"
strquery = "Select*from issued_books"
rs.Open strquery, con, adOpenDynamic, adLockOptimistic
rs.MoveFirst
While Not rs.EOF
If Val(txtbid.Text) = rs.Fields("bid") Then
rs.Delete
End If
rs.MoveNext
Wend
rs.Close
con.Close


con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data source=" & App.Path & "\clgdata.mdb;Persist Security Info=false"
strquery = "Select*from books"
rs.Open strquery, con, adOpenDynamic, adLockOptimistic
rs.MoveFirst
While Not rs.EOF
If rs!bid = txtbid Then
 rs!bookstatus = "Available"
 rs.Update
 MsgBox ("Book Returned Successfully")
End If
rs.MoveNext
Wend
rs.Close
con.Close

End If
End Sub

Private Sub btncancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Height = 0
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
If Me.Height < 7770 Then
Me.Height = Me.Height + 200
Else
Timer1.Enabled = False
End If
End Sub


Private Sub txtbid_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
If txtbid = "" Then
MsgBox ("Enter Book ID & Press Enter")
Else

con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data source=" & App.Path & "\clgdata.mdb;Persist Security Info=false"
strquery = "Select*from issued_books"
rs.Open strquery, con, adOpenDynamic, adLockOptimistic
Dim flag As Boolean
flag = False

rs.MoveFirst
While Not rs.EOF
If Val(txtbid.Text) = rs.Fields("bid") Then
 txtbname.Text = rs.Fields("bname")
 txtmname.Text = rs.Fields("mname")
 txtmemberid.Text = rs.Fields("memberid")
 DTPicker1.Value = rs.Fields("issue_date")
 DTPicker2.Value = Date
 flag = True
End If
rs.MoveNext
Wend
If flag = False Then
MsgBox ("Book Not Found In Issued Record")
End If

Dim d1, d2, d3, fine As Integer
fine = 0
d1 = DTPicker1.Value
d2 = DTPicker2.Value
d3 = d2 - d1
If d3 > 15 Then
fine = 50
End If
txtfine.Text = fine

rs.Close
con.Close
End If

End If
End Sub

