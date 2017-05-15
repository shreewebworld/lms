VERSION 5.00
Begin VB.Form login 
   BackColor       =   &H0080C0FF&
   Caption         =   "Login"
   ClientHeight    =   3000
   ClientLeft      =   7875
   ClientTop       =   4155
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   6615
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   4440
      Picture         =   "frmlogin.frx":0000
      ScaleHeight     =   1875
      ScaleWidth      =   1755
      TabIndex        =   6
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000080FF&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Geometr212 BkCn BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Geometr212 BkCn BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Adodc2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Geometr212 BkCn BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "UserID"
      BeginProperty Font 
         Name            =   "Geometr212 BkCn BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim strquery As String

Private Sub Command1_Click()
If Text1.Text = rs!userid And Text2.Text = rs!Password Then
MsgBox ("Login success")
MDIForm1.mnulogin1.Enabled = False
'MDIForm1.Show
'MDIForm1.Enabled = True
'MDIForm1.mnubooks.Enabled = True
'MDIForm1.mnumembers.Enabled = True
'MDIForm1.mnureport.Enabled = True
'MDIForm1.mnuhelp.Enabled = True
frmmain.Command3.Visible = True
frmmain.Command4.Visible = True
frmmain.Command5.Visible = True
frmmain.Frame1.Visible = False
Unload Me
Else
MsgBox ("Invalid UserID or Password ")
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data source=" & App.Path & "\clgdata.mdb;Persist Security Info=False"
strquery = "Select*from users"
rs.Open strquery, con, adOpenDynamic, adLockOptimistic
End Sub

Private Sub Form_Unload(Cancel As Integer)
rs.Close
con.Close
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Text1.Text = rs!userid And Text2.Text = rs!Password Then
MsgBox ("login success")
Unload Me
MDIForm1.Show
MDIForm1.Enabled = True
MDIForm1.mnubooks.Enabled = True
MDIForm1.mnumembers.Enabled = True
MDIForm1.mnureport.Enabled = True
MDIForm1.mnuhelp.Enabled = True
Else
MsgBox ("Invalid UserID or Password ")
End If
End If
End Sub
