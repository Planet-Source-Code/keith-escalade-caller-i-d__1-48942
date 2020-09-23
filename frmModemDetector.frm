VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form ModemDetect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modem detector -- Must be offline"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   435
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
   Icon            =   "frmModemDetector.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin MSCommLib.MSComm COM1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InputLen        =   1
      RThreshold      =   1
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2040
      MaxLength       =   2
      TabIndex        =   5
      Text            =   "3"
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtModem 
      BackColor       =   &H000080FF&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2640
      Width           =   4815
   End
   Begin VB.CommandButton btnNext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Next >"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton btnDetect 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Detect"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin VB.ListBox lstTasks 
      BackColor       =   &H000080FF&
      ForeColor       =   &H00FFFFFF&
      Height          =   2400
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Modem name:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Scan COM port 1 to"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   1815
   End
End
Attribute VB_Name = "ModemDetect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FullModem As String
Dim FullPort As Long
Private Sub SearchPort(sPort As Long, whatCom As MSComm)
Dim sInput As String
On Error Resume Next
whatCom.PortOpen = False
whatCom.CommPort = sPort
whatCom.Settings = "9600,N,8,1"
whatCom.InputLen = 0
whatCom.PortOpen = True
whatCom.Output = "ATI4" & Chr(13)
sInput = whatCom.Input
Debug.Print sInput
whatCom.PortOpen = False
If sInput = "" Then UpdateStatus "Searched COM" & sPort: UpdateStatus Error: UpdateStatus "": Exit Sub
If sInput <> "" Then UpdateStatus "Searched COM" & sPort: UpdateStatus "Modem Found": UpdateStatus Mid(sInput, 3, Len(sInput) - 8): UpdateStatus "": txtModem.Text = Mid(sInput, 3, Len(sInput) - 8): LogSettings Mid(sInput, 3, Len(sInput) - 8), sPort: FullModem = Mid(sInput, 3, Len(sInput) - 8): FullPort = sPort: btnNext.Enabled = True: Exit Sub
End Sub
Private Sub btnDetect_Click()
If Text1 < 1 Then MsgBox "Number must be and integer higher than zero", vbInformation: Exit Sub
txtModem.Text = ""
btnDetect.Enabled = False
btnNext.Enabled = False
Dim X As Long
lstTasks.Clear
For X = 1 To Val(Text1.Text)
SearchPort X, COM1
DoEvents
Next X
DoEvents
btnDetect.Enabled = True
End Sub
Private Sub UpdateStatus(sText As String)
lstTasks.AddItem sText
End Sub
Private Sub btnNext_Click()
BallerID.Text2.Text = FullModem
BallerID.Text3.Text = FullPort
BallerID.Show
Me.Hide
End Sub
Private Sub Form_activate()
On Error GoTo 3
Dim sText$
Close #1
Open App.Path & "/Baller_ID_Modem_Settings.ini" For Input As #1
While Not EOF(1) = True
Input #1, sText$
If Left(sText$, 13) = "[Modem Name] " Then themodem = Right(sText$, Len(sText$) - 13)
If Left(sText$, 13) = "[Modem Port] " Then theport = Right(sText$, Len(sText$) - 13)
Wend
Close #1
If theport < 1 Or theport > 99 Then Exit Sub
Dim a As String
a = MsgBox("Would you like to skip the modem detection stage?", vbYesNo)
If a = vbNo Then Exit Sub
BallerID.Text2.Text = themodem
BallerID.Text3.Text = Val(theport)
BallerID.Show
Me.Hide
3 End Sub
Private Sub Form_Unload(Cancel As Integer)
End
End Sub
Private Sub LogSettings(ModemName As String, sPort As Long)
Close #1
Open App.Path & "/Baller_ID_Modem_Settings.ini" For Output As #1
Print #1, "[Baller_ID_Modem_Settings]" & vbCrLf & "[Modem Name] " & ModemName & vbCrLf & "[Modem Port] " & sPort
Close #1
End Sub
