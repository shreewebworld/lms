VERSION 5.00
Begin VB.Form frmmain 
   Caption         =   "Library Management System"
   ClientHeight    =   7050
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   13425
   LinkTopic       =   "Form1"
   Picture         =   "frmmain.frx":0000
   ScaleHeight     =   7050
   ScaleWidth      =   13425
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Library Management System"
      BeginProperty Font 
         Name            =   "Castellar"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   2040
      TabIndex        =   0
      Top             =   840
      Width           =   13245
   End
   Begin VB.Menu mnulogin 
      Caption         =   "Login"
      Begin VB.Menu mnulogin1 
         Caption         =   "Log-in"
      End
      Begin VB.Menu mnulogout 
         Caption         =   "Logout"
      End
      Begin VB.Menu a 
         Caption         =   "-"
      End
      Begin VB.Menu mnuadduser 
         Caption         =   "Add New User"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Hide
login.Show
login.SetFocus

End Sub

