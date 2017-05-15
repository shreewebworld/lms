VERSION 5.00
Begin VB.Form frmnewuser 
   BackColor       =   &H00C0FFC0&
   Caption         =   "New User Registration"
   ClientHeight    =   4305
   ClientLeft      =   7110
   ClientTop       =   3690
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   ScaleHeight     =   4305
   ScaleWidth      =   7965
   Begin VB.TextBox txtuname 
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
      Left            =   2880
      TabIndex        =   6
      Top             =   2640
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   5400
      Picture         =   "frmnewuser.frx":0000
      ScaleHeight     =   1875
      ScaleWidth      =   1875
      TabIndex        =   9
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000080FF&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "Create User"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox txtpass1 
      DataSource      =   "Adodc1"
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
      IMEMode         =   3  'DISABLE
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox txtpass 
      DataSource      =   "Adodc1"
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
      IMEMode         =   3  'DISABLE
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox txtuserid 
      DataSource      =   "Adodc1"
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
      Left            =   2880
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "User's Name"
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
      Left            =   1200
      TabIndex        =   10
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm-Password"
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
      Left            =   840
      TabIndex        =   2
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Left            =   1680
      TabIndex        =   1
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
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
      Left            =   1920
      TabIndex        =   0
      Top             =   840
      Width           =   855
   End
End
Attribute VB_Name = "frmnewuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim strquery As String
Private Sub Command1_Click()
If txtpass.Text = txtpass1.Text Then
rs.AddNew
rs.Fields("userid") = txtuserid.Text
rs.Fields("password") = txtpass.Text
rs.Fields("uname") = txtuname.Text
rs.Update
MsgBox ("New User Created Successfully !!!")
Else
MsgBox ("Password does not match")
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data source=" & App.Path & "\clgdata.mdb;Persist Security Info=false"
strquery = "Select*from users"
rs.Open strquery, con, adOpenDynamic, adLockOptimistic
End Sub

Private Sub Form_Unload(Cancel As Integer)
rs.Close
con.Close
End Sub

