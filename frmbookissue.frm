VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmbookissue 
   BackColor       =   &H00000040&
   Caption         =   "Book Issue"
   ClientHeight    =   7860
   ClientLeft      =   7605
   ClientTop       =   2070
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   ScaleHeight     =   7860
   ScaleWidth      =   6690
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2760
      Top             =   0
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2760
      TabIndex        =   17
      Text            =   "Combo1"
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox txtcategory 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   16
      Top             =   4320
      Width           =   1815
   End
   Begin VB.TextBox txtauther 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   14
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000040&
      Caption         =   "Enter BookID and Press Enter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   960
      TabIndex        =   10
      Top             =   1440
      Width           =   4815
      Begin VB.TextBox txtbid 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1800
         TabIndex        =   11
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Book ID"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   480
         TabIndex        =   12
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.TextBox txtbname 
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
      Left            =   2760
      TabIndex        =   9
      Top             =   3120
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   6120
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
      Format          =   107216897
      CurrentDate     =   42223
   End
   Begin VB.TextBox txtmname 
      BackColor       =   &H00FFFFFF&
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
      Left            =   2760
      TabIndex        =   7
      Top             =   5520
      Width           =   2655
   End
   Begin VB.CommandButton btncancel 
      BackColor       =   &H000080FF&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton btnissue 
      BackColor       =   &H0000FF00&
      Caption         =   "Issue"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   960
      Picture         =   "frmbookissue.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Book Category"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   840
      TabIndex        =   15
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label lblauther 
      BackStyle       =   0  'Transparent
      Caption         =   "Auther"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1680
      TabIndex        =   13
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Member ID"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1200
      TabIndex        =   6
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Member Name"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Book Issue"
      BeginProperty Font 
         Name            =   "Colonna MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Issue Date"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Book Name"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   3120
      Width           =   1335
   End
End
Attribute VB_Name = "frmbookissue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim strquery As String

Private Sub btncancel_Click()
Unload Me
End Sub

Private Sub btnissue_Click()
If txtbid.Text = "" Then
MsgBox ("Please enter book id & press enter")
ElseIf txtmname.Text = "" Then
MsgBox ("Select Member ID & Press Enter")
Else

con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data source=" & App.Path & "\clgdata.mdb;Persist Security Info=false"
strquery = "Select * from issued_books"
rs.Open strquery, con, adOpenDynamic, adLockOptimistic
rs.AddNew
rs!bname = txtbname.Text
rs!bid = Val(txtbid.Text)
rs!mname = txtmname.Text
rs!memberid = Val(Combo1.Text)
rs!issue_date = DTPicker1.Value
rs.Update
rs.Close
con.Close
btncancel.Caption = "Close"

con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data source=" & App.Path & "\clgdata.mdb;Persist Security Info=false"
strquery = "Select * from transactionn"
rs.Open strquery, con, adOpenDynamic, adLockOptimistic
rs.AddNew
rs!bname = txtbname.Text
rs!bid = Val(txtbid.Text)
rs!mname = txtmname.Text
rs!memberid = Val(Combo1.Text)
rs!issue_date = DTPicker1.Value
rs.Update
rs.Close
con.Close

con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data source=" & App.Path & "\clgdata.mdb;Persist Security Info=false"
strquery = "Select * from books"
rs.Open strquery, con, adOpenDynamic, adLockOptimistic
rs.MoveFirst
While Not rs.EOF
If rs!bid = txtbid Then
 rs!bookstatus = "Issued"
 rs.Update
 MsgBox ("Book Issued Successfully")
End If
rs.MoveNext
Wend
rs.Close
con.Close

txtbid.Text = ""
txtauther.Text = ""
txtbname.Text = ""
txtcategory.Text = ""
txtmname.Text = ""
Combo1.Text = ""


End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

flag = False
Dim cnt As Integer

con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data source=" & App.Path & "\clgdata.mdb;Persist Security Info=false"
strquery = "Select * from issued_books"
rs.Open strquery, con, adOpenDynamic, adLockOptimistic
rs.MoveFirst
While Not rs.EOF
If rs!memberid = Combo1.SelText Then
cnt = cnt + 1
End If
rs.MoveNext
Wend
rs.Close
con.Close

If cnt >= 2 Then
MsgBox ("User Already Issued " & cnt & " Books." & vbNewLine & " Please Return Book & Try Again Later.")
txtmname.Text = ""
Else

con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data source=" & App.Path & "\clgdata.mdb;Persist Security Info=false"
strquery = "Select * from members"
rs.Open strquery, con, adOpenDynamic, adLockOptimistic

rs.MoveFirst
While Not rs.EOF
If rs!memberid = Combo1.Text Then
txtmname.Text = rs!mname
Dim a As Integer
a = rs!memberid

flag = True
End If

rs.MoveNext
Wend
If flag = False Then
MsgBox ("Enter Correct Member ID")
End If
rs.Close
con.Close

End If
End If
End Sub

Private Sub Form_Activate()
Me.Height = 0
Timer1.Enabled = True
txtbid.SetFocus
End Sub

Private Sub Form_Load()

con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data source=" & App.Path & "\clgdata.mdb;Persist Security Info=false"
strquery = "Select * from members"
rs.Open strquery, con, adOpenDynamic, adLockOptimistic
rs.MoveFirst
While Not rs.EOF
Combo1.AddItem (rs!memberid)

rs.MoveNext
Wend

DTPicker1.Value = Date
rs.Close
con.Close
End Sub

Private Sub Timer1_Timer()
If Me.Height < 8085 Then
Me.Height = Me.Height + 200
Else
Timer1.Enabled = False
End If
End Sub

Private Sub txtbid_KeyPress(KeyAscii As Integer)

con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data source=" & App.Path & "\clgdata.mdb;Persist Security Info=false"
strquery = "Select * from books"
rs.Open strquery, con, adOpenDynamic, adLockOptimistic

Dim flag As Boolean
flag = False
If KeyAscii = 13 Then
rs.MoveFirst
While Not rs.EOF
If rs!bid = txtbid.Text Then


flag = True
 If rs!bookstatus = "Issued" Then
 MsgBox ("Book AllReady Issued")
 Else
 txtbname.Text = rs!bname
txtauther.Text = rs!auther
txtcategory.Text = rs!category
 End If

End If
rs.MoveNext
Wend
If flag = False Then
MsgBox "Book Not Available"
End If
End If

rs.Close
con.Close
End Sub

