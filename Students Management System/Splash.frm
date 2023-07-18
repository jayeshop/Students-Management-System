VERSION 5.00
Begin VB.Form Splash 
   BackColor       =   &H00C0C000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   Picture         =   "Splash.frx":0000
   ScaleHeight     =   6015
   ScaleWidth      =   9585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   7320
      Top             =   6000
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Timer1.Enabled = False
  Unload Me
  Login.Show
End Sub
