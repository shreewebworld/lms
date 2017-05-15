VERSION 5.00
Begin VB.Form frmaddmember 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Add New Member"
   ClientHeight    =   5430
   ClientLeft      =   7275
   ClientTop       =   2400
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   5865
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   2280
      TabIndex        =   12
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   240
      Top             =   4320
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000080FF&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Corbel"
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
      TabIndex        =   9
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   3600
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   3000
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   2400
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Member ID"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   11
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   945
      Left            =   600
      Picture         =   "frmaddmember.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   915
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "New Member Registration"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   495
      Left            =   1560
      TabIndex        =   10
      Top             =   360
      Width           =   6255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No."
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Member Name"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   1800
      Width           =   1575
   End
End
Attribute VB_Name = "frmaddmember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim strquery As String

Private Sub Command1_Click()

If Text1.Text = "" Then
MsgBox ("Please Enter Name")
Text1.SetFocus
ElseIf IsNumeric(Text1.Text) Then
MsgBox ("Invalid Name")
ElseIf Text4.Text = "" Then
MsgBox ("Please Enter Adddress")
Text4.SetFocus
Else

con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data source=" & App.Path & "\clgdata.mdb;Persist Security Info=false"
strquery = "Select*from members"
rs.Open strquery, con, adOpenDynamic, adLockOptimistic


rs.AddNew
rs.Fields("memberid") = Text2.Text
rs.Fields("mname") = Text1.Text
rs.Fields("memail") = Text3.Text
rs.Fields("mphone") = Text4.Text
rs.Fields("maddr") = Text5.Text
rs.Update

Command2.Caption = "Close"
MsgBox ("Member Added Successfully !!!")
rs.Close
con.Close

Text1.Text = ""
Text2.Text = Val(Text2.Text) + 1
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""



End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Height = 0
Timer1.Enabled = True
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data source=" & App.Path & "\clgdata.mdb;Persist Security Info=false"
strquery = "Select*from members"
rs.Open strquery, con, adOpenDynamic, adLockOptimistic
Dim max As Integer
rs.MoveFirst
While Not rs.EOF
If max < rs!memberid Then
max = rs!memberid
End If
rs.MoveNext
Wend
Text2.Text = max + 1
rs.Close
con.Close
End Sub

Private Sub Timer1_Timer()
If Me.Height < 6345 Then
Me.Height = Me.Height + 200
Else
Timer1.Enabled = False
End If
End Sub
