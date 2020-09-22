VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Creating New Archive"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8745
   Icon            =   "frmNew.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   8745
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5400
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   4995
      TabIndex        =   7
      Top             =   3120
      Width           =   5055
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Create"
      Enabled         =   0   'False
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
      Left            =   4440
      TabIndex        =   6
      ToolTipText     =   "Click Here to Pack all files in the List"
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton Help 
      Caption         =   "Help"
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton cmdCls 
      Caption         =   "Clear List"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Browse"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   0
      Width           =   855
   End
   Begin VB.ListBox List2 
      Height          =   2205
      Left            =   9000
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   3015
      Left            =   5400
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCls_Click()
Dim Re
If List1.ListCount = 0 Then Exit Sub
Re = MsgBox("Are You Sure want to Clear List", vbCritical + vbYesNo, "Confirmation")
If Re = vbYes Then
List1.Clear
List2.Clear
End If
If List1.ListCount = 0 Then
cmdRemove.Enabled = False
cmdOk.Enabled = False
cmdCls.Enabled = False
Else
cmdRemove.Enabled = True
cmdOk.Enabled = True
cmdCls.Enabled = True
End If
End Sub

Private Sub cmdOk_Click()
Dim Con
If List1.ListCount = 0 Then Exit Sub
If frmMain.NewCr = True Then
 Con = MsgBox("Sure want to Create New Archive", vbInformation + vbYesNo, "Confirmation")
 If Con = vbYes Then

 CommonDialog1.Filter = "Pic-Packed Files (.PPK)|*.ppk"
 CommonDialog1.ShowSave
 If Len(CommonDialog1.FileName) Then CreateArchiv (CommonDialog1.FileName)
Unload Me
 End If
 Else
 Con = MsgBox("Sure want to Add File to Archive", vbInformation + vbYesNo, "Confirmation")
 If Con = vbYes Then
 AddF frmMain.OpendArch, List1.ListCount
 OpenArchiv frmMain.OpendArch
 Unload Me
 End If
 End If
End Sub

Private Sub cmdRemove_Click()
If List1.Text <> "" Then
List1.RemoveItem (List1.ListIndex)
List2.RemoveItem (List2.ListIndex)
Else
If List1.ListCount = 0 Then MsgBox "No Item to Remove", vbCritical, "Pic-Packer": Exit Sub
MsgBox "Please Select a Item to Remove From List", vbCritical, "Pic-Packer"
End If
If List1.ListCount = 0 Then
cmdRemove.Enabled = False
cmdOk.Enabled = False
cmdCls.Enabled = False
Else
cmdRemove.Enabled = True
cmdOk.Enabled = True
cmdCls.Enabled = True
End If
End Sub

Private Sub cmdAdd_Click()
frmAdd.Show 1
If List1.ListCount = 0 Then
cmdRemove.Enabled = False
cmdOk.Enabled = False
cmdCls.Enabled = False
Else
cmdRemove.Enabled = True
cmdOk.Enabled = True
cmdCls.Enabled = True
End If
End Sub





Private Sub Form_Load()
If frmMain.NewCr = True Then
Me.Caption = "Creating New Archive"
Else
Me.Caption = "Adding File into ('" & frmMain.OpendArch & "')"
cmdOk.Caption = "Add"
End If
End Sub

Private Sub List1_Click()
On Error Resume Next
Dim F As String
List2.Selected(List1.ListIndex) = True
Me.Caption = "Creating New Archive " & List1.ListCount & "/" & List1.ListIndex + 1
F = List1.Text
Image1.Picture = LoadPicture(F)

If Error <> "" Then MsgBox "Error : " & Error
End Sub

