VERSION 5.00
Begin VB.Form frmdeletebook 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Delete Book"
   ClientHeight    =   5970
   ClientLeft      =   6885
   ClientTop       =   2790
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   5145
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Enter BookID and Press Enter"
      ForeColor       =   &H00008000&
      Height          =   1215
      Left            =   720
      TabIndex        =   9
      Top             =   1320
      Width           =   3735
      Begin VB.TextBox txtbid 
         BackColor       =   &H0080FFFF&
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
         Left            =   1440
         TabIndex        =   11
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Book Id"
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
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.CommandButton btndelete 
      BackColor       =   &H000080FF&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton btnnext 
      BackColor       =   &H0080FF80&
      Caption         =   "Next  >"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton btnprev 
      BackColor       =   &H0080FF80&
      Caption         =   "<  Prev"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox txtauther 
      DataField       =   "auther"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   4440
      Width           =   1935
   End
   Begin VB.TextBox txtcategory 
      DataField       =   "bid"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox txtbname 
      DataField       =   "bname"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   2280
      TabIndex        =   3
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Book Delete"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1440
      TabIndex        =   12
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Auther"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Book Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   3120
      Width           =   1215
   End
End
Attribute VB_Name = "frmdeletebook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim strquery As String


Private Sub btnnext_Click()
rs.MoveNext
If rs.EOF Then
rs.MoveFirst
End If
txtbname.Text = rs!bname
txtcategory.Text = rs.Fields("category")
txtauther.Text = rs.Fields("auther")
End Sub


Private Sub btndelete_Click()
Dim ans As Integer
ans = MsgBox("Really want to delete Book :" & rs!bname, vbYesNo)
If ans = vbYes Then
rs.Delete
MsgBox ("Book Deleted")
End If
End Sub

Private Sub btnprev_Click()
rs.MovePrevious
If rs.BOF Then
rs.MoveLast
End If
txtbname.Text = rs.Fields("bname")
txtcategory.Text = rs.Fields("category")
txtauther.Text = rs.Fields("auther")
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data source=" & App.Path & "\clgdata.mdb;Persist Security Info=false"
strquery = "Select*from books"

rs.Open strquery, con, adOpenDynamic, adLockOptimistic

txtbname.Text = rs.Fields("bname")
txtauther.Text = rs.Fields("auther")
txtcategory.Text = rs.Fields("category")
End Sub

Private Sub Form_Unload(Cancel As Integer)
rs.Close
con.Close
End Sub


Private Sub txtbid_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

Dim flag As Boolean
flag = False

rs.MoveFirst
While Not rs.EOF
If rs!bid = txtbid.Text Then
txtbname.Text = rs.Fields("bname")
txtcategory.Text = rs.Fields("category")
txtauther.Text = rs.Fields("auther")
ans = MsgBox("Really Want to Delete This Book :" & rs!bname, vbYesNo, "Delete")
If ans = vbYes Then
rs.Delete
MsgBox (txtbname.Text & " Book Deleted ")
End If
flag = True
End If

rs.MoveNext
Wend
If flag = False Then
MsgBox ("Book Not Found")
End If
rs.MoveFirst
End If
End Sub

