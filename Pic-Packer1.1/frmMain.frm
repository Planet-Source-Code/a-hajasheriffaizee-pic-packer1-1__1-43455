VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pic-Packer         By A. HajaSherifFaizee"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   6585
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5280
      TabIndex        =   8
      ToolTipText     =   "Add More Pics in Opened Archive"
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4320
      TabIndex        =   7
      ToolTipText     =   "Delete Selected Pic"
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extract"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      ToolTipText     =   "Extract Selected Pic"
      Top             =   2520
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      ScaleHeight     =   195
      ScaleWidth      =   3555
      TabIndex        =   5
      Top             =   3840
      Width           =   3615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4560
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      ToolTipText     =   "Open an Existing Archive"
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Create New Archive"
      Top             =   3600
      Width           =   735
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   0
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      ToolTipText     =   "Dobule Click to View"
      Top             =   0
      Width           =   3375
   End
   Begin VB.Label Disply 
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   2535
      Left            =   3480
      Stretch         =   -1  'True
      ToolTipText     =   "Dobule Click to View"
      Top             =   0
      Width           =   3015
   End
   Begin VB.Menu mnuExtract 
      Caption         =   "Extract"
      Visible         =   0   'False
      Begin VB.Menu ExtractOne 
         Caption         =   "Extract Selected File"
      End
      Begin VB.Menu mnuExAll 
         Caption         =   "Extract All"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public OpendArch As String
Dim T As String
Public NewCr As Boolean

Private Sub cmdAdd_Click()
NewCr = False
frmNew.Show 1
End Sub

Private Sub cmdDel_Click()
Dim Del As Boolean
Dim L As Long
Dim Con As String
If List1.SelCount = 0 Then Exit Sub
If List1.SelCount > 1 Then
Con = MsgBox("Sure want to Delete " & List1.SelCount & " Files from ('" & OpendArch & "')", vbCritical + vbYesNo)
Else
Con = MsgBox("Sure want to Delete ('" & List1.Text & "') from ('" & OpendArch & "')", vbCritical + vbYesNo)
End If
If Con = vbYes Then
DeleteF List1.SelCount, OpendArch
OpenArchiv OpendArch
End If
End Sub

Private Sub cmdExtract_Click()
Dim i As Long
If List1.SelCount > 0 Then
Picture1.Refresh
If List1.SelCount = 1 Then
ExtractOne.Caption = "Extract ('" & List1.Text & "')"
Else
ExtractOne.Caption = "Extract " & List1.SelCount & " Files "
End If
PopupMenu mnuExtract
Else
List1.Selected(0) = True
cmdExtract_Click
End If
ExtractOne.Caption = "Extract Selected File"
End Sub

Private Sub cmdNew_Click()
NewCr = True
frmNew.Show 1
End Sub

Private Sub cmdOpen_Click()
CommonDialog1.Filter = "Pic-Packed Files (.PPK)|*.ppk"
CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Then Exit Sub
If Dir(CommonDialog1.FileName) <> "" Then
 T = OpenArchiv(CommonDialog1.FileName)
OpendArch = CommonDialog1.FileName
Me.Caption = "Pic-Packer (' " & OpendArch & " ')"
Disply.Caption = "Total : " & T
End If
End Sub

Private Sub Command1_Click()
End
End Sub



Private Sub Image1_DblClick()
If List1.Text = "" Then Exit Sub
frmView.Show 1
End Sub

Private Sub List1_DblClick()
If List1.ListCount = 0 Then Exit Sub
On Error Resume Next
ViewArchiv List1.ListIndex
Disply.Caption = "Total : " & OPendTotal & "  Size : " & FSize & " bytes"

If Error <> "" Then MsgBox "Error : " & Error
End Sub

Private Sub mnuExAll_Click()
Dim R As String
Dim Fol As String
Fol = BrowseForFolder("Select a Folder to Extract All files")
If Len(Fol) Then
If Dir(Fol, vbDirectory + vbNormal) <> "" Then R = Extract(True, Fol): MsgBox R & " Files Extracted SuccessFully"
End If
End Sub

Private Sub ExtractOne_Click()
Dim Fol As String
Dim R As String
Fol = BrowseForFolder("Select a Folder to Extract Selected File")
If Len(Fol) Then
If Dir(Fol, vbDirectory + vbNormal) <> "" Then R = Extract(False, Fol): MsgBox R & " Files Extracted SuccessFully"
End If
End Sub

