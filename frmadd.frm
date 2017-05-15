VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmaddbook 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Add New Book"
   ClientHeight    =   6615
   ClientLeft      =   6885
   ClientTop       =   2985
   ClientWidth     =   5700
   Icon            =   "frmadd.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   5700
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2280
      Top             =   0
   End
   Begin VB.TextBox txtprice 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3000
      TabIndex        =   10
      Top             =   4200
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "frmadd.frx":119A2
      Left            =   2520
      List            =   "frmadd.frx":119B8
      TabIndex        =   11
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000080FF&
      Caption         =   "Cancel"
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
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "ADD"
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
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5880
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      DataField       =   "date"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   5280
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
      Format          =   106692609
      CurrentDate     =   42209
   End
   Begin VB.TextBox txtpublication 
      DataField       =   "publication"
      DataSource      =   "Adodc1"
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
      Left            =   2520
      TabIndex        =   9
      Top             =   3600
      Width           =   2655
   End
   Begin VB.TextBox txtauther 
      DataField       =   "auther"
      DataSource      =   "Adodc1"
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
      Left            =   2520
      TabIndex        =   8
      Top             =   3000
      Width           =   2775
   End
   Begin VB.TextBox txtbid 
      DataField       =   "bid"
      DataSource      =   "Adodc1"
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
      Left            =   2520
      TabIndex        =   7
      Top             =   2400
      Width           =   2775
   End
   Begin VB.TextBox txtbname 
      DataField       =   "bname"
      DataSource      =   "Adodc1"
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
      Left            =   2520
      TabIndex        =   6
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Rs."
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
      Left            =   2640
      TabIndex        =   17
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
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
      Left            =   1440
      TabIndex        =   16
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Add Book"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   15
      Top             =   720
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   1200
      Left            =   1080
      Picture         =   "frmadd.frx":119EA
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1185
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
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
      Left            =   1560
      TabIndex        =   5
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
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
      TabIndex        =   4
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Publication"
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
      Left            =   1080
      TabIndex        =   3
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Auther"
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
      Left            =   1440
      TabIndex        =   2
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Book ID"
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
      TabIndex        =   1
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name Of Book"
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
      Left            =   840
      TabIndex        =   0
      Top             =   1800
      Width           =   1335
   End
End
Attribute VB_Name = "frmaddbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim strquery As String
Private Sub Command1_Click()
If txtbname.Text = "" Then
MsgBox ("Please Enter Book Name")
txtbname.SetFocus
ElseIf txtbid = "" Then
MsgBox ("Please Enter Book ID ")
txtbid.SetFocus
ElseIf txtauther = "" Then
MsgBox ("Please Enter Auther")
txtauther.SetFocus
ElseIf Combo1.Text = "" Then
MsgBox ("Please Select Category")
Combo1.SetFocus
ElseIf Not IsNumeric(txtprice.Text) Then
MsgBox ("Invalid Price")
txtprice.SetFocus
ElseIf Not IsNumeric(txtbid.Text) Then
MsgBox ("Invalid Book ID")
txtbid.SetFocus
Else

con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data source=" & App.Path & "\clgdata.mdb;Persist Security Info=false"
strquery = "Select*from books"
rs.Open strquery, con, adOpenDynamic, adLockOptimistic

rs.AddNew
rs!bname = txtbname.Text
rs!bid = txtbid
rs!auther = txtauther.Text
rs!publication = txtpublication.Text
rs!category = Combo1.Text
rs!price = Val(txtprice.Text)
rs!Date = DTPicker1.Value
rs!bookstatus = "Available"
rs.Update
MsgBox ("Book Successfully Added !!!")

rs.Close
con.Close

txtbname.Text = ""
txtbid.Text = ""
txtauther.Text = ""
txtpublication.Text = ""
txtprice.Text = ""
Combo1.Text = ""

Command2.Caption = "Close"

End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Height = 0
Timer1.Enabled = True

' txtbname.Text = ""
'txtbid.Text = ""
'txtauther.Text = ""
'txtpublication.Text = ""
Combo1.Text = ""
DTPicker1.Value = Date

End Sub


Private Sub Timer1_Timer()
If Me.Height < 7260 Then
Me.Height = Me.Height + 200
Else
Timer1.Enabled = False
End If
End Sub
