VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmupdatebook 
   Caption         =   "Update Books"
   ClientHeight    =   5190
   ClientLeft      =   7275
   ClientTop       =   3180
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   Picture         =   "frmupdate.frx":0000
   ScaleHeight     =   5190
   ScaleWidth      =   5460
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFF00&
      Caption         =   "Next "
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FF80FF&
      Caption         =   "Prev"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000080FF&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3960
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      DataField       =   "date"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   2880
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   106692609
      CurrentDate     =   42209
   End
   Begin VB.TextBox txtpublication 
      DataField       =   "publication"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox txtauther 
      DataField       =   "auther"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtbid 
      DataField       =   "bid"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txtbname 
      DataField       =   "bname"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Publication"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
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
      ForeColor       =   &H00FF00FF&
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Book ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   840
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
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmupdatebook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim strquery As String
Private Sub Command1_Click()
rs!bname = txtbname.Text
rs!auther = txtauther.Text
rs!publication = txtpublication.Text
rs!Date = DTPicker1.Value
rs.Update
MsgBox ("Book Updated Successfully !!!")
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
rs.MovePrevious
If rs.BOF Then
rs.MoveLast
End If
txtbname.Text = rs.Fields("bname")
txtbid.Text = rs.Fields("bid")
txtauther.Text = rs.Fields("auther")
txtpublication.Text = rs.Fields("publication")
DTPicker1.Value = rs.Fields("date")
End Sub

Private Sub Command4_Click()
rs.MoveNext
If rs.EOF Then
rs.MoveFirst
End If
txtbname.Text = rs.Fields("bname")
txtbid.Text = rs.Fields("bid")
txtauther.Text = rs.Fields("auther")
txtpublication.Text = rs.Fields("publication")
DTPicker1.Value = rs.Fields("date")
End Sub

Private Sub Form_Load()

con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data source=" & App.Path & "\clgdata.mdb;Persist Security Info=false"
strquery = "Select*from books"
rs.Open strquery, con, adOpenDynamic, adLockOptimistic

End Sub

Private Sub Form_Unload(Cancel As Integer)
rs.Close
con.Close
End Sub
