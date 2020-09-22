VERSION 5.00
Begin VB.Form frmAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adding File(s) into an Archive"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8385
   Icon            =   "frmAdd.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   8385
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Cancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6720
      TabIndex        =   6
      Top             =   3240
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmAdd.frx":030A
      Left            =   2640
      List            =   "frmAdd.frx":0320
      Style           =   2  'Dropdown List
      TabIndex        =   5
      ToolTipText     =   "Pattern"
      Top             =   3360
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add All Files"
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Selected Files"
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   3240
      Width           =   1335
   End
   Begin VB.FileListBox File1 
      Height          =   3210
      Left            =   2400
      MultiSelect     =   1  'Simple
      Pattern         =   "*.bmp;*.jpg;*.gif;*.ico;*.jpeg"
      TabIndex        =   2
      Top             =   0
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   2295
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   3135
      Left            =   5160
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim L
Dim a As Boolean
Dim FileP As String


Private Sub Cancel_Click()
Unload Me
End Sub

Private Sub Combo1_Change()
File1.Pattern = Combo1.Text
File1.Refresh
End Sub

Private Sub Combo1_Click()
File1.Pattern = Combo1.Text
File1.Refresh
End Sub

Private Sub Combo1_DblClick()
File1.Pattern = Combo1.Text
File1.Refresh
End Sub

Private Sub Command1_Click()
If Right(Dir1.Path, 1) <> "\" Then
FileP = Dir1.Path & "\"
Else
FileP = Dir1.Path
End If
For L = 0 To File1.ListCount - 1
Check (File1.List(L))
If File1.Selected(L) = True Then
If a = False Then
a = True
frmNew.List1.AddItem (FileP & File1.List(L))
frmNew.List2.AddItem (File1.List(L))
End If
End If
Next L
Unload Me
End Sub

Private Sub Command2_Click()
If Right(Dir1.Path, 1) <> "\" Then
FileP = Dir1.Path & "\"
Else
FileP = Dir1.Path
End If
For L = 0 To File1.ListCount - 1
Check (File1.List(L))
If a = False Then
a = True
frmNew.List1.AddItem (FileP & File1.List(L))
frmNew.List2.AddItem (File1.List(L))
End If
Next L
Unload Me
End Sub

Private Sub Dir1_Change()
On Error Resume Next
File1.Path = Dir1
If Error <> "" Then MsgBox "Error : " & Error, vbCritical, "Error :"
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1
If Error <> "" Then MsgBox "Error : " & Error, vbCritical, "Error :"
End Sub

Function Check(File As String) As Boolean
Dim O
If frmNew.List2.ListCount = 0 Then a = False
For O = 0 To frmNew.List2.ListCount - 1
If File = frmNew.List2.List(O) Then
a = True
Exit Function
Else
a = False
End If
Next O
End Function





Private Sub File1_Click()
Dim F As String
If Right(Dir1.Path, 1) <> "\" Then
F = Dir1.Path & "\"
Else
F = Dir1.Path
End If
Image1.Picture = LoadPicture(F & File1)
End Sub

Private Sub Form_Load()
Drive1.Refresh
Dir1.Refresh
File1.Path = Dir1
File1.Refresh
End Sub
