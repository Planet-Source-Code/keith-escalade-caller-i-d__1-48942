VERSION 5.00
Object = "{EEE78583-FE22-11D0-8BEF-0060081841DE}#1.0#0"; "XVoice.dll"
Begin VB.Form Speaker 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Speaker"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8115
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSpeaker.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   8115
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVEVOICEPROJECTLibCtl.DirectSS DirectS 
      Height          =   375
      Left            =   0
      OleObjectBlob   =   "frmSpeaker.frx":0CCE
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H000080FF&
      ForeColor       =   &H00FFFFFF&
      Height          =   2565
      Left            =   4680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "frmSpeaker.frx":0D26
      Top             =   120
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Test"
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ok"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
   Begin VB.ListBox lstSpeaker 
      BackColor       =   &H000080FF&
      ForeColor       =   &H00FFFFFF&
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Speaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Visible = False
End Sub
Private Sub Command2_Click()
DirectS.Speak Text1.Text
End Sub
Private Sub Command3_Click()
On Error Resume Next
Close #1
BallerID.DirectS.Select (lstSpeaker.ListIndex) + 1
Open App.Path & "/Baller_ID_Speaker.ini" For Output As #1
Print #1, "Speaker=" & lstSpeaker.ListIndex
Close #1
Me.Visible = False
End Sub
Private Sub Form_Load()
Dim X As Long
For X = 1 To DirectS.CountEngines
lstSpeaker.AddItem DirectS.ModeName(X)
Next X
End Sub
Private Sub Form_Unload(Cancel As Integer)
Cancel = True
Me.Visible = False
End Sub
Private Sub lstSpeaker_Click()
On Error Resume Next
DirectS.Select (lstSpeaker.ListIndex) + 1
End Sub
