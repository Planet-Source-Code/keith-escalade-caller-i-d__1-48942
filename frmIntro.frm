VERSION 5.00
Begin VB.Form Intro 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Baller I.D."
   ClientHeight    =   3210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9975
   Icon            =   "frmIntro.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3210
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   960
      Top             =   600
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   720
      Top             =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2460
      Left            =   0
      Picture         =   "frmIntro.frx":0CCE
      ScaleHeight     =   2430
      ScaleWidth      =   7680
      TabIndex        =   0
      Top             =   0
      Width           =   7710
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Skip(3)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   1
         Top             =   2040
         Width           =   855
      End
   End
End
Attribute VB_Name = "Intro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Show
Me.Width = Picture1.Width
Me.Height = Picture1.Height
Me.Top = Screen.Height / 2 - Me.Height / 2
Me.Left = Screen.Width / 2 - Me.Width / 2
End Sub
Private Sub Label1_Click()
Timer1.Enabled = False
Timer2.Enabled = False
ModemDetect.Show
Unload Me
End Sub
Private Sub Timer1_Timer()
Do Until Me.Height < 100
Me.Height = Me.Height - 100
Me.Top = Screen.Height / 2 - Me.Height / 2
Me.Left = Screen.Width / 2 - Me.Width / 2
Pause 0.0001
DoEvents
Loop
Do Until Me.Width > Screen.Width
Me.Width = Me.Width + 200
DoEvents
Me.Left = Me.Left - 100
DoEvents
Pause 0.0001
Loop
ModemDetect.Show
Me.Hide
Timer1.Enabled = False
End Sub
Sub Pause(interval)
current = Timer
Do While Timer - current < Val(interval)
DoEvents
Loop
End Sub
Private Sub Timer2_Timer()
If Label1.Caption = "Skip(3)" Then Label1.Caption = "Skip(2)": Exit Sub
If Label1.Caption = "Skip(2)" Then Label1.Caption = "Skip(1)": Exit Sub
If Label1.Caption = "Skip(1)" Then Label1.Caption = "Skip(0)": Exit Sub
If Label1.Caption = "Skip(0)" Then Timer2.Enabled = False: Exit Sub
End Sub
