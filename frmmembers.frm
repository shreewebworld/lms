VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmmembersreport 
   BackColor       =   &H0080FF80&
   Caption         =   "Members"
   ClientHeight    =   5490
   ClientLeft      =   7275
   ClientTop       =   4500
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   9765
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
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   5318
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
      Height          =   810
      Left            =   360
      Picture         =   "frmmembers.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Members Report"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   570
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   3075
   End
End
Attribute VB_Name = "frmmembersreport"
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
strquery = "Select*from members"
rs.Open strquery, con, adOpenDynamic, adLockOptimistic
MSFlexGrid1.Cols = 4
MSFlexGrid1.TextMatrix(0, 0) = "Name"
MSFlexGrid1.TextMatrix(0, 1) = "M ID"
MSFlexGrid1.TextMatrix(0, 2) = "Email"
MSFlexGrid1.TextMatrix(0, 3) = "Address"
rs.MoveFirst
r = 1
Do While Not rs.EOF
MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
MSFlexGrid1.TextMatrix(r, 0) = rs!mname
MSFlexGrid1.TextMatrix(r, 1) = rs!memberid
MSFlexGrid1.TextMatrix(r, 2) = rs!memail
MSFlexGrid1.TextMatrix(r, 3) = rs!maddr
rs.MoveNext
r = r + 1
Loop

MSFlexGrid1.ColWidth(0) = 1500
MSFlexGrid1.ColWidth(1) = 500
MSFlexGrid1.ColWidth(2) = 2500
MSFlexGrid1.ColWidth(3) = 2500
End Sub

Private Sub Form_Unload(Cancel As Integer)
rs.Close
con.Close
End Sub
