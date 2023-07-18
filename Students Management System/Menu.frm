VERSION 5.00
Begin VB.Form Menu 
   BackColor       =   &H00C0C000&
   Caption         =   "Menu"
   ClientHeight    =   4095
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7110
   BeginProperty Font 
      Name            =   "Modern No. 20"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Menu.frx":0000
   ScaleHeight     =   4095
   ScaleWidth      =   7110
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Edit Student Details"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   4
      Top             =   2280
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add New Student"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   2280
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Student Information"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   2
      Top             =   1080
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LOGOUT"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   7080
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7080
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "STUDENTS MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Login.Show
Unload Me
End Sub

Private Sub Command2_Click()
stdInfo.Show
Unload Me
End Sub

Private Sub Command3_Click()
NewStd.Show
Unload Me
End Sub

Private Sub Command4_Click()
EditStd.Show
Unload Me
End Sub
