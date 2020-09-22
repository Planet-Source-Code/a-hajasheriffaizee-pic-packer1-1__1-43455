VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmView 
   BackColor       =   &H00000000&
   Caption         =   "Viewing"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   8370
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3480
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Left            =   2880
      Top             =   240
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmView.frx":0000
      Left            =   1560
      List            =   "frmView.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdAuto 
      Caption         =   "Slide Show"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "Previous"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   2880
      ScaleHeight     =   3015
      ScaleWidth      =   4575
      TabIndex        =   0
      Top             =   2880
      Width           =   4575
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cx As Integer
Dim cy As Integer
Dim cw As Integer
Dim ch As Integer
Dim x As Integer
Dim y As Integer
Dim Total As Long
Dim Index As Long

Private Sub cmdAuto_Click()
Index = 1
Timer1.Interval = 2000
Timer1.Enabled = True
End Sub

Private Sub cmdNext_Click()
On Error Resume Next
If Index = OPendTotal - 1 Then MsgBox "Noting to View": Exit Sub

Index = Index + 1
ViewArchiv Index
Picture1.Picture = frmMain.Image1.Picture
CenterPic Picture1, 0, 0, Screen.Width, Screen.Height
Me.Caption = "Viewing ('" & frmMain.List1.List(Index) & "')"
If Error <> "" Then MsgBox "Error : " & Error
End Sub

Private Sub cmdPrevious_Click()
On Error Resume Next
If Index = 0 Then MsgBox "Noting to View": Exit Sub
Index = Index - 1
ViewArchiv Index
Picture1.Picture = frmMain.Image1.Picture
CenterPic Picture1, 0, 0, Screen.Width, Screen.Height
Me.Caption = "Viewing ('" & frmMain.List1.List(Index) & "')"
If Error <> "" Then MsgBox "Error : " & Error
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
With CommonDialog1
.DialogTitle = "Picture Save as "
.Filter = "*.*|*.*"
.ShowSave
If Len(.FileName) Then SavePicture Picture1, .FileName & ".bmp"
End With
If Error <> "" Then MsgBox "Error : " & Error
End Sub

Private Sub Combo1_Change()
If Combo1.ListIndex = 0 Then Timer1.Interval = 1000
If Combo1.ListIndex = 1 Then Timer1.Interval = 2000
If Combo1.ListIndex = 2 Then Timer1.Interval = 3000
If Combo1.ListIndex = 3 Then Timer1.Interval = 4000
End Sub

Private Sub Combo1_Click()
If Combo1.ListIndex = 0 Then Timer1.Interval = 1000
If Combo1.ListIndex = 1 Then Timer1.Interval = 2000
If Combo1.ListIndex = 2 Then Timer1.Interval = 3000
If Combo1.ListIndex = 3 Then Timer1.Interval = 4000
End Sub

Private Sub Form_Load()
On Error Resume Next
Picture1.Picture = frmMain.Image1.Picture
CenterPic Picture1, 0, 0, Screen.Width, Screen.Height
Index = 1
If Error <> "" Then MsgBox "Error : " & Error
End Sub
Public Sub CenterPic(Nam As PictureBox, sx, sy, swidth, sheight)
 cx = sx + swidth / 2
 cy = sy + sheight / 2
 cw = Nam.Width / 2
 ch = Nam.Height / 2
 x = cx - cw
 y = cy - ch
 Nam.Left = x
 Nam.Top = y
End Sub

Private Sub Timer1_Timer()
ViewArchiv Index
Picture1.Picture = frmMain.Image1.Picture
CenterPic Picture1, 0, 0, Screen.Width, Screen.Height
Me.Caption = "Viewing ('" & frmMain.List1.List(Index) & "')"
Index = Index + 1
If Index = OPendTotal - 1 Then MsgBox "Done !!!": Timer1.Enabled = False
End Sub
