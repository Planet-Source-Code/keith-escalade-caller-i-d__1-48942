VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Pics 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pictures..."
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6495
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPictures.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   6495
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ok"
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3840
      Width           =   1215
   End
   Begin VB.ListBox List4 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5520
      TabIndex        =   18
      Top             =   4080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H000080FF&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   2640
      Width           =   5175
   End
   Begin VB.ListBox List3 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5400
      TabIndex        =   16
      Top             =   3960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5280
      TabIndex        =   15
      Top             =   3840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Apply"
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Remove"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton btnAdd 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Add"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "..."
      Height          =   255
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H000080FF&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   3360
      Width           =   4575
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H000080FF&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   3000
      Width           =   5175
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H000080FF&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   2280
      Width           =   5175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Names"
      ForeColor       =   &H00000000&
      Height          =   2055
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3495
      Begin VB.ListBox lstNames 
         BackColor       =   &H000080FF&
         ForeColor       =   &H00FFFFFF&
         Height          =   1620
         ItemData        =   "frmPictures.frx":0CCE
         Left            =   120
         List            =   "frmPictures.frx":0CD0
         TabIndex        =   7
         Top             =   240
         Width           =   3255
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   6000
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Picture Preview"
      ForeColor       =   &H00000000&
      Height          =   2055
      Left            =   3720
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H000080FF&
         Height          =   1695
         Left            =   120
         ScaleHeight     =   1635
         ScaleWidth      =   2355
         TabIndex        =   5
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "CID Name :"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Picture File :"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "CID Number :"
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   975
   End
End
Attribute VB_Name = "Pics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAdd_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Then MsgBox "Please Fill in All Information.", vbInformation: Exit Sub
Dim X As Long
For X = 0 To lstNames.ListCount - 1
If LCase(lstNames.List(X)) = LCase(Text1.Text) Then MsgBox "Name already added to list", vbInformation: Exit Sub
Next X
Dim xu As Long
For xu = 0 To lstNames.ListCount - 1
If LCase(lstNames.List(xu)) = LCase(Text3.Text) Then MsgBox "CID Number already added to list", vbInformation: Exit Sub
Next xu
Dim xi As Long
For xi = 0 To lstNames.ListCount - 1
If LCase(lstNames.List(xi)) = LCase(Text2.Text) Then MsgBox "CID Name already added to list", vbInformation: Exit Sub
Next xi
lstNames.AddItem Text1.Text
List2.AddItem Text2.Text
List3.AddItem Text3.Text
List4.AddItem Text4.Text
SaveAll
End Sub
Private Sub Command1_Click()
On Error GoTo 3
CD1.Filter = "Picture Files|*.jpg;*.jpeg;*.gif;*.bmp|"
CD1.CancelError = True
CD1.ShowOpen
Text4.Text = CD1.FileName
Picture1.Picture = LoadPicture(CD1.FileName)
3 End Sub
Private Sub Command2_Click()
On Error Resume Next
lstNames.RemoveItem lstNames.ListIndex
List2.RemoveItem List2.ListIndex
List3.RemoveItem List3.ListIndex
List4.RemoveItem List4.ListIndex
End Sub
Private Sub Command3_Click()
On Error GoTo 3
Dim X As Long
For X = 0 To lstNames.ListCount - 1
If LCase(lstNames.List(X)) = LCase(Text1.Text) Then MsgBox "Name already added to list", vbInformation: Exit Sub
Next X
Dim xu As Long
For xu = 0 To lstNames.ListCount - 1
If LCase(lstNames.List(xu)) = LCase(Text3.Text) Then MsgBox "CID Number already added to list", vbInformation: Exit Sub
Next xu
Dim xi As Long
For xi = 0 To lstNames.ListCount - 1
If LCase(lstNames.List(xi)) = LCase(Text2.Text) Then MsgBox "CID Name already added to list", vbInformation: Exit Sub
Next xi
If lstNames.ListIndex < 0 Then Exit Sub
lstNames.List(lstNames.ListIndex) = Text1.Text
List2.List(List2.ListIndex) = Text2.Text
List3.List(List3.ListIndex) = Text3.Text
List4.List(List4.ListIndex) = Text4.Text
3 End Sub
Private Sub Command4_Click()
Me.Hide
End Sub
Private Sub Form_Load()
LoadFiles
End Sub
Private Sub Form_Unload(Cancel As Integer)
Cancel = True
Me.Hide
End Sub
Private Sub Label1_Click()
MsgBox "Examples: 1231234567, 1237654321, 1230123456", vbInformation
End Sub
Private Sub lstNames_Click()
List2.ListIndex = lstNames.ListIndex
List3.ListIndex = lstNames.ListIndex
List4.ListIndex = lstNames.ListIndex
Text1.Text = lstNames.List(lstNames.ListIndex)
Text2.Text = List2.List(List2.ListIndex)
Text3.Text = List3.List(List3.ListIndex)
Text4.Text = List4.List(List4.ListIndex)
If Text4.Text <> "" Then Picture1.Picture = LoadPicture(Text4.Text)
End Sub
Sub SaveList(sList As ListBox, sInfo$)
Dim X As Long
Open App.Path & "/Baller_ID_" & sInfo & ".ini" For Output As #1
For X = 0 To sList.ListCount - 1
Print #1, sList.List(X)
Next X
DoEvents
Close #1
End Sub
Sub SaveAll()
Close #1
SaveList lstNames, "Pic_Names"
DoEvents
SaveList List2, "CID_Names"
DoEvents
SaveList List3, "CID_Numbers"
DoEvents
SaveList List4, "Picture_Files"
End Sub
Sub LoadFiles()
On Error GoTo 3
Close #1
Open App.Path & "/Baller_ID_Pic_Names.ini" For Input As #1
While Not EOF(1) = True
Input #1, sText$
lstNames.AddItem sText$
DoEvents
Wend
Close #1
Open App.Path & "/Baller_ID_CID_Names.ini" For Input As #1
While Not EOF(1) = True
Input #1, sText$
List2.AddItem sText$
DoEvents
Wend
Close #1
Open App.Path & "/Baller_ID_CID_Numbers.ini" For Input As #1
While Not EOF(1) = True
Input #1, sText$
List3.AddItem sText$
DoEvents
Wend
Close #1
Open App.Path & "/Baller_ID_Picture_Files.ini" For Input As #1
While Not EOF(1) = True
Input #1, sText$
List4.AddItem sText$
DoEvents
Wend
Close #1
3 End Sub
