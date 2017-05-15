VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmbookreport 
   BackColor       =   &H0080FF80&
   Caption         =   "Books Report"
   ClientHeight    =   4875
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12300
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   12300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3735
      Left            =   3000
      TabIndex        =   0
      Top             =   1080
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   6588
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   4875
      Left            =   0
      Picture         =   "frmbookreport.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3000
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Books Report"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   495
      Left            =   4200
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "frmbookreport"
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
strquery = "Select*from books"
rs.Open strquery, con, adOpenDynamic, adLockOptimistic
MSFlexGrid1.Cols = 6
MSFlexGrid1.TextMatrix(0, 0) = "ID"
MSFlexGrid1.TextMatrix(0, 1) = "Book Name"
MSFlexGrid1.TextMatrix(0, 2) = "Auther"
MSFlexGrid1.TextMatrix(0, 3) = "Publication"
MSFlexGrid1.TextMatrix(0, 4) = "Price"
MSFlexGrid1.TextMatrix(0, 5) = "Status"
rs.MoveFirst
r = 1
Do While Not rs.EOF
MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
MSFlexGrid1.TextMatrix(r, 0) = rs!bid
MSFlexGrid1.TextMatrix(r, 1) = rs!bname
MSFlexGrid1.TextMatrix(r, 2) = rs!auther
MSFlexGrid1.TextMatrix(r, 3) = rs!publication
MSFlexGrid1.TextMatrix(r, 4) = rs!price
MSFlexGrid1.TextMatrix(r, 5) = rs!bookstatus
rs.MoveNext
r = r + 1
Loop

MSFlexGrid1.ColWidth(0) = 500
MSFlexGrid1.ColWidth(1) = 2000
MSFlexGrid1.ColWidth(2) = 2000
MSFlexGrid1.ColWidth(3) = 2000
End Sub

Private Sub Form_Unload(Cancel As Integer)
rs.Close
con.Close
End Sub
