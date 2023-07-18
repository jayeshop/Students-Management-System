VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form EditStd 
   BackColor       =   &H00C0C000&
   Caption         =   "Edit Student Details"
   ClientHeight    =   10125
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21585
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
   Picture         =   "EditStd.frx":0000
   ScaleHeight     =   10125
   ScaleWidth      =   21585
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "MENU"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   19200
      TabIndex        =   23
      Top             =   9480
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
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
      Height          =   495
      Left            =   16800
      TabIndex        =   22
      Top             =   9480
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "EditStd.frx":21FB7
      Height          =   4095
      Left            =   120
      TabIndex        =   8
      Top             =   4920
      Width           =   21255
      _ExtentX        =   37491
      _ExtentY        =   7223
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16777152
      ForeColor       =   4210688
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Students Information"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc stdInfo 
      Height          =   375
      Left            =   16440
      Top             =   240
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\BBSC Project\New\Database\Student Information.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\BBSC Project\New\Database\Student Information.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SudentsInformation"
      Caption         =   "stdInfo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text6 
      DataField       =   "Email Address"
      DataSource      =   "stdInfo"
      Height          =   405
      Left            =   12000
      TabIndex        =   7
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      DataField       =   "Mobile Number"
      DataSource      =   "stdInfo"
      Height          =   405
      Left            =   12000
      TabIndex        =   6
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      DataField       =   "Previous Institution"
      DataSource      =   "stdInfo"
      Height          =   495
      Left            =   7080
      TabIndex        =   5
      Top             =   3360
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      DataField       =   "Address"
      DataSource      =   "stdInfo"
      Height          =   735
      Left            =   7080
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2520
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      DataField       =   "Date of Birth"
      DataSource      =   "stdInfo"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   3240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      _Version        =   393216
      Format          =   127729665
      CurrentDate     =   36526
   End
   Begin VB.TextBox Text2 
      DataField       =   "Mother Name"
      DataSource      =   "stdInfo"
      Height          =   405
      Left            =   2640
      TabIndex        =   2
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      DataField       =   "Name"
      DataSource      =   "stdInfo"
      Height          =   405
      Left            =   2640
      TabIndex        =   1
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Edit Details"
      Height          =   3375
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   21135
      Begin VB.ComboBox Combo3 
         DataField       =   "Year"
         DataSource      =   "stdInfo"
         Height          =   390
         ItemData        =   "EditStd.frx":21FCD
         Left            =   15720
         List            =   "EditStd.frx":21FDA
         TabIndex        =   26
         Text            =   "Select"
         Top             =   2280
         Width           =   1695
      End
      Begin VB.ComboBox Combo2 
         DataField       =   "Course"
         DataSource      =   "stdInfo"
         Height          =   390
         ItemData        =   "EditStd.frx":22003
         Left            =   11880
         List            =   "EditStd.frx":22016
         TabIndex        =   25
         Text            =   "Select"
         Top             =   2280
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "Gender"
         DataSource      =   "stdInfo"
         Height          =   390
         ItemData        =   "EditStd.frx":22035
         Left            =   6960
         List            =   "EditStd.frx":2203F
         TabIndex        =   24
         Text            =   "Select"
         Top             =   720
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "DELETE"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   18000
         TabIndex        =   21
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "SAVE"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   15480
         TabIndex        =   20
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "Year :"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   14040
         TabIndex        =   19
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "Course :"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   9840
         TabIndex        =   18
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "Email Address :"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   9840
         TabIndex        =   17
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "Mobile No :"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   9840
         TabIndex        =   16
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "Previous Institution :"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   4440
         TabIndex        =   15
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "Address :"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   5400
         TabIndex        =   14
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "Gender :"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   5280
         TabIndex        =   13
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "Date of Birth :"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   600
         TabIndex        =   12
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "Mother Name :"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   480
         TabIndex        =   11
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "Name :"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   21600
      Y1              =   9240
      Y2              =   9240
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   21600
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   21600
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "STUDENTS MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      TabIndex        =   0
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "EditStd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
stdInfo.Recordset.Update
End Sub

Private Sub Command2_Click()
answer = MsgBox("Do you want to delete this record?", vbExclamation + vbYesNo, "Confirm")
If answer = vbYes Then
    stdInfo.Recordset.Delete
Else
    MsgBox "Cancelled", vbInformation, "Confirm"
End If

End Sub

Private Sub Command3_Click()
Login.Show
Unload Me
End Sub

Private Sub Command4_Click()
Menu.Show
Unload Me
End Sub
