VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmissuedbooksreport 
   BackColor       =   &H0000FF00&
   Caption         =   "Form1"
   ClientHeight    =   5130
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   7125
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "Exit"
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
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   4895
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   1080
      Picture         =   "frmissuedbooksreport.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Issued Books"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   360
      Width           =   3135
   End
End
Attribute VB_Name = "frmissuedbooksreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim strquery As String

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data source=" & App.Path & "\clgdata.mdb;Persist Security Info=false"
strquery = "Select*from issued_books"
rs.Open strquery, con, adOpenDynamic, adLockOptimistic
MSFlexGrid1.Cols = 5
MSFlexGrid1.TextMatrix(0, 0) = "B ID"
MSFlexGrid1.TextMatrix(0, 1) = "Book Name"
MSFlexGrid1.TextMatrix(0, 2) = "M Id"
MSFlexGrid1.TextMatrix(0, 3) = "M Name"
MSFlexGrid1.TextMatrix(0, 4) = "Issue Date"

rs.MoveFirst
r = 1
Do While Not rs.EOF
MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
MSFlexGrid1.TextMatrix(r, 0) = rs!bid
MSFlexGrid1.TextMatrix(r, 1) = rs!bname
MSFlexGrid1.TextMatrix(r, 2) = rs!memberid
MSFlexGrid1.TextMatrix(r, 3) = rs!mname
MSFlexGrid1.TextMatrix(r, 4) = rs!issue_date

rs.MoveNext
r = r + 1
Loop

MSFlexGrid1.ColWidth(0) = 500
MSFlexGrid1.ColWidth(1) = 1500
MSFlexGrid1.ColWidth(2) = 500
MSFlexGrid1.ColWidth(3) = 1500
MSFlexGrid1.ColWidth(3) = 1500
End Sub

Private Sub Form_Unload(Cancel As Integer)
rs.Close
con.Close
End Sub
