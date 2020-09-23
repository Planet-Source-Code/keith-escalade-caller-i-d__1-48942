VERSION 5.00
Begin VB.Form PopUp 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPop-Up.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2295
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   7000
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Caller's Picture"
      ForeColor       =   &H00000000&
      Height          =   2055
      Left            =   3600
      TabIndex        =   9
      Top             =   120
      Width           =   2655
      Begin VB.PictureBox Picture1 
         BackColor       =   &H000080FF&
         Height          =   1695
         Left            =   120
         ScaleHeight     =   1635
         ScaleWidth      =   2355
         TabIndex        =   10
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Caller's Identification"
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3375
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label lblNumb 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Time :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Numb :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   3
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Name :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exit"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   3375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Incoming Call From"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "PopUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Timer1.Enabled = True
If Me.Visible = True Then Exit Sub
Me.Top = Screen.Height / 2 - Me.Height / 2
Me.Left = Screen.Width - Me.Width
Me.Left = Screen.Width - Me.Width
Me.lblName = BallerID.PCalling.Text
Me.lblNumb = BallerID.PNumber.Text
Me.lblTime = Time
End Sub
Private Sub Label5_Click()
Me.Visible = False
End Sub
Private Sub Timer1_Timer()
Unload Me
Timer1.Enabled = False
End Sub
