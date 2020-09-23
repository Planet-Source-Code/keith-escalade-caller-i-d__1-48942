VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{EEE78583-FE22-11D0-8BEF-0060081841DE}#1.0#0"; "XVoice.dll"
Begin VB.Form BallerID 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Baller I.D."
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10215
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBallerID.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   10215
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Text4 
      BackColor       =   &H000080FF&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmBallerID.frx":0CCE
      Left            =   1560
      List            =   "frmBallerID.frx":0CEA
      TabIndex        =   23
      Text            =   "AT+VCID=1"
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox PNumber 
      Height          =   285
      Left            =   3360
      TabIndex        =   21
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox PCalling 
      Height          =   285
      Left            =   3240
      TabIndex        =   20
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   6720
      Top             =   360
   End
   Begin ACTIVEVOICEPROJECTLibCtl.DirectSS DirectS 
      Height          =   375
      Left            =   6000
      OleObjectBlob   =   "frmBallerID.frx":0D3E
      TabIndex        =   17
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6600
      Top             =   240
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Caller's Picture"
      ForeColor       =   &H00000000&
      Height          =   2055
      Left            =   7440
      TabIndex        =   14
      Top             =   0
      Width           =   2655
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H000080FF&
         Height          =   1695
         Left            =   120
         ScaleHeight     =   1635
         ScaleWidth      =   2355
         TabIndex        =   15
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Stats"
      ForeColor       =   &H00000000&
      Height          =   4935
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   9975
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Last Incoming Call (Click name to configure pictures)"
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   9735
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   9495
         End
      End
      Begin VB.ListBox lstTasks 
         BackColor       =   &H000080FF&
         ForeColor       =   &H00FFFFFF&
         Height          =   3765
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   9735
      End
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Show Pop-Up on Incoming Call"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H000080FF&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5520
      LinkTimeout     =   0
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   7
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H000080FF&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1560
      LinkTimeout     =   0
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   5
      Top             =   1800
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H000080FF&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3120
      TabIndex        =   2
      Text            =   "call from"
      Top             =   360
      Width           =   975
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Speak Identification"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listen for Incoming Calls"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin MSCommLib.MSComm COM1 
      Left            =   5280
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Enable CID String:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7080
      TabIndex        =   19
      Top             =   7200
      Width           =   3015
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Speaker..."
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Pictures..."
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Not Monitoring Calls"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00000000&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   240
      Shape           =   3  'Circle
      Top             =   7200
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Modem Port:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4200
      TabIndex        =   6
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Modem Name:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Pre-Text:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   360
      Width           =   975
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "mnuMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "Show"
      End
      Begin VB.Menu mnuPictures 
         Caption         =   "Pictures"
      End
      Begin VB.Menu mnuSpeaker 
         Caption         =   "Speaker"
      End
      Begin VB.Menu mnuConnectModem 
         Caption         =   "Connect Modem"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "BallerID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This has not been fully tested out. Please report all bugs/errors to keith_escalade@yahoo.com
' keith_escalade 2003
Dim RinSec As Long
Public PersonCalling, PersonsNumber As String
Private Sub Check1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorHandle
Select Case Check1.Value
Case Is = 1
mnuConnectModem.Caption = "Disconnect Modem"
COM1.CommPort = Text3.Text
COM1.PortOpen = True ' Open port and monitor calls
COM1.Output = Text4.Text & vbCr ' Signal the modem to look for caller identification
Shape1.BackColor = vbGreen
Shape1.FillColor = vbGreen
Label4.Caption = "Monitoring Calls"
Case Is = 0
mnuConnectModem.Caption = "Connect Modem"
COM1.PortOpen = False
Shape1.BackColor = vbRed
Shape1.FillColor = vbRed
Label4.Caption = "Not Monitoring Calls"
End Select
Exit Sub
ErrorHandle:
MsgBox Err.Description, vbCritical, "Error"
mnuConnectModem.Caption = "Connect Modem"
Check1.Value = 0
End Sub
Private Sub DialNumber(sNumber As Long, sPort As Long, whatCom As MSComm)
whatCom.CommPort = sPort ' Select port
whatCom.PortOpen = True ' Open port
whatCom.Output = "ATDT " & sNumber & vbCr ' "AT"tention "D"ial "T"one
' whatCom.output = "ATDP " & sNumber & vbcr < Pulse dialing
' "AT"tention "D"ial "P"ulse
End Sub
Private Sub COM1_OnComm()
On Error Resume Next
Dim cInput$
cInput = COM1.Input ' Stores the input from com port
Debug.Print cInput
' Get Name and Number
If InStr(1, cInput, "NMBR = ") Then
PersonCalling = GetItem(cInput, "NAME = ")
PersonsNumber = GetItem(cInput, "NMBR = ")
PersonCalling = Right(PersonCalling, Len(PersonCalling) - 7)
PersonsNumber = Right(PersonsNumber, Len(PersonsNumber) - 7)
If PersonCalling = "O" Then PersonCalling = "Unknown Name" ' Returns Unknown if caller is 'O'.
If PersonsNumber = "O" Then PersonsNumber = "Unknown Number" ' Returns Unknown if caller is 'O'.
Dim xu As Long
For xu = 0 To Pics.List2.ListCount - 1
If LCase(PersonCalling) = LCase(Pics.List2.List(xu)) Then If LCase(PersonsNumber) = LCase(Pics.List3.List(xu)) Then PersonCalling = Pics.lstNames.List(xu): Picture1.Picture = LoadPicture(Pics.List4.List(xu)): PopUp.Picture1.Picture = LoadPicture(Pics.List4.List(xu))
DoEvents
Next xu
PCalling.Text = PersonCalling ' Store caller for other forms
PNumber.Text = PersonsNumber ' Store number for other forms
If Check3.Value = 1 Then PopUp.Visible = True: PopUp.Timer1.Enabled = True
If Check2.Value = 1 Then DirectS.Speak Text1.Text & " " & PersonsNumber & " " & PersonCalling
Label5.Caption = Format(PersonsNumber, "(###) ###-####") & " " & PersonCalling ' Format number of person calling
AddTask "CALL " & Format(PersonsNumber, "(###) ###-####") & " " & PersonCalling & " " & Now
SaveTasks
End If
Select Case COM1.CommEvent
Case OnCommConstants.comEvRing ' On phone ring event
Label4.Caption = "Ringing"
AddTask "RING " & Now
SaveTasks
Timer1.Enabled = True
RinSec = 0
End Select
End Sub
Private Sub AddTask(sTask As String)
lstTasks.AddItem sTask
End Sub
Sub AnswerPhone(whatCom As MSComm)
' "AT"tention "A"nswer
whatCom.Output = "ATA" & vbCr
End Sub
Sub HangUp(whatCom As MSComm)
' "AT"tention "H"angup
whatCom.Output = "ATH" & vbCr
End Sub
Private Sub Form_Initialize()
sIcon.cbSize = Len(sIcon)
sIcon.hwnd = Me.hwnd
sIcon.uId = vbNull
sIcon.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
sIcon.uCallBackMessage = WM_MOUSEMOVE
sIcon.hIcon = Me.Icon
sIcon.szTip = "Baller I.D." & vbNullChar
Call Shell_NotifyIcon(NIM_ADD, sIcon)
Call Shell_NotifyIcon(NIM_MODIFY, sIcon)
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case X
Case 7755: 'Right click icon event
PopupMenu mnuMenu
Case 7725:
Me.Show 'Double click icon event
End Select
End Sub
Private Sub Form_Load()
On Error Resume Next
Dim sText1, sText2, sText3, sText4 As String
Label8.Caption = Now
Appends
Close #1
Open App.Path & "/Baller_ID_PreText.ini" For Input As #1
Input #1, sText1
Text1.Text = sText1
Close #1
Open App.Path & "/Baller_ID_CID_Enabled.ini" For Input As #1
If EOF(1) = True Then GoTo asdf
Input #1, sText4
Text4.Text = sText4
asdf:
Close #1
Open App.Path & "/Baller_ID_Tasks.ini" For Input As #1
While EOF(1) = False
Input #1, sText2
lstTasks.AddItem sText2
DoEvents
Wend
Close #1
Open App.Path & "/Baller_ID_Speaker.ini" For Input As #1
Input #1, sText3
On Error Resume Next
Speaker.lstSpeaker.ListIndex = Right(sText3, Len(sText3) - 8) - 1
If Left(sText3, 8) = "Speaker=" Then DirectS.Select Right(sText3, Len(sText3) - 8): Speaker.DirectS.Select Right(sText3, Len(sText3) - 8)
End Sub
Sub Appends()
' Creates a blank file if not already made
Close #1
Open App.Path & "/Baller_ID_PreText.ini" For Append As #1
Close #1
Open App.Path & "/Baller_ID_Tasks.ini" For Append As #1
Close #1
Open App.Path & "/Baller_ID_Speaker.ini" For Append As #1
Close #1
Open App.Path & "/Baller_ID_Pic_Names.ini" For Append As #1
Close #1
Open App.Path & "/Baller_ID_CID_Names.ini" For Append As #1
Close #1
Open App.Path & "/Baller_ID_CID_Numbers.ini" For Append As #1
Close #1
Open App.Path & "/Baller_ID_Picture_Files.ini" For Append As #1
Close #1
Open App.Path & "/Baller_ID_CID_Enabled.ini" For Append As #1
Close #1
End Sub
Private Sub Form_Unload(Cancel As Integer)
Close #1
Open App.Path & "/Baller_ID_CID_Enabled.ini" For Output As #1
Print #1, Text4.Text
Close #1
DoEvents
Open App.Path & "/Baller_ID_PreText.ini" For Output As #1
Print #1, Text1.Text ' Save pre-text
Close #1
DoEvents
Cancel = True
Me.Hide
End Sub
Private Sub Label5_Click()
If Label5.Caption = "" Then MsgBox "No Name Displayed.", vbInformation: Exit Sub
Pics.Text1.Text = ""
Pics.Text2.Text = PersonCalling
Pics.Text3.Text = PersonsNumber
Pics.Text4.Text = ""
Pics.Show
End Sub
Private Sub Label6_Click()
Pics.Show
End Sub
Private Sub Label7_Click()
Speaker.Visible = True
End Sub
Private Sub Timer1_Timer()
' Enabled when phone is ringing
Dim X As Long
Shape1.BackColor = vbBlue
Shape1.FillColor = vbBlue
Pause 0.5
Shape1.BackColor = &HFF8080
Shape1.FillColor = &HFF8080
RingAgain:
RinSec = 0
For X = 0 To 7
DoEvents
RinSec = RinSec + 1
If RinSec = 7 Then If Check1.Value = 1 Then Shape1.BackColor = vbGreen: Shape1.FillColor = vbGreen: Label4.Caption = "Monitoring Calls": Timer1.Enabled = False: Exit Sub
If RinSec = 7 Then If Check1.Value = 0 Then Shape1.BackColor = vbRed: Shape1.FillColor = vbRed: Label4.Caption = "Not Monitoring Calls": Timer1.Enabled = False: Exit Sub
If RinSec = 0 Then GoTo RingAgain
Pause 1#
Next X
End Sub
Function GetItem(sData As String, sExcludedData As String)
' Parse caller identification
Dim start1, end1, product As String
start1 = InStr(1, sData, sExcludedData)
end1 = InStr(start1 + 1, sData, vbCrLf)
product = Mid(sData, start1, end1 - start1)
GetItem = product
End Function
Sub Pause(interval)
Dim current$
current = Timer
Do While Timer - current < Val(interval)
DoEvents
Loop
End Sub
Private Sub Timer2_Timer()
Label8.Caption = Now
End Sub
Sub mnuShow_Click()
Me.Show
End Sub
Sub mnuExit_click()
End
End Sub
Sub mnuSpeaker_click()
Label7_Click
End Sub
Sub mnuPictures_click()
Label6_Click
End Sub
Sub mnuConnectModem_click()
Select Case mnuConnectModem.Caption
Case Is = "Disconnect Modem"
mnuConnectModem.Caption = "Connect Modem"
Check1.Value = 0
COM1.PortOpen = False
Shape1.BackColor = vbRed
Shape1.FillColor = vbRed
Label4.Caption = "Not Monitoring Calls"
Case Is = "Connect Modem"
On Error GoTo ErrorHandler
mnuConnectModem.Caption = "Disconnect Modem"
Check1.Value = 1
COM1.CommPort = Text3.Text
COM1.PortOpen = True
COM1.Output = Text4.Text & vbCr
Shape1.BackColor = vbGreen
Shape1.FillColor = vbGreen
Label4.Caption = "Monitoring Calls"
End Select
Exit Sub
ErrorHandler:
MsgBox Err.Description, vbCritical, "Error"
mnuConnectModem.Caption = "Connect Modem"
Check1.Value = 0
End Sub
Sub SaveTasks()
Close #1
Open App.Path & "/Baller_ID_Tasks.ini" For Output As #1
Dim xi As Long
For xi = 0 To lstTasks.ListCount - 1
Print #1, lstTasks.List(xi)
Next xi
DoEvents
Close #1
End Sub
