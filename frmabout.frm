VERSION 5.00
Begin VB.Form frmabout 
   BackColor       =   &H00000000&
   Caption         =   "About"
   ClientHeight    =   8235
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   10875
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image2 
      Height          =   3735
      Left            =   7800
      Picture         =   "frmabout.frx":0000
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Thanks"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   4080
      TabIndex        =   9
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   3180
      Left            =   2400
      Picture         =   "frmabout.frx":10356
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   5805
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Talegaon Dhamdhere Tal - Shirur Dist - Pune"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3360
      TabIndex        =   8
      Top             =   7440
      Width           =   6495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "S.S. Dhamdhere College Of Arts And Commerce"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   2520
      TabIndex        =   7
      Top             =   6840
      Width           =   7575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "TYBCA"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7200
      TabIndex        =   6
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Kishor Jedhe"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   7080
      TabIndex        =   5
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Dnyaneshwar Naiknavare"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   5520
      Width           =   3495
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Developed By :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "TYBCA"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "This Project is  Created to Automate Functions of Lirary Using Computer System."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   840
      Width           =   8415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Library Management System"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Width           =   4305
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


