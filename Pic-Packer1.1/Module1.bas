Attribute VB_Name = "ModMain"
Option Explicit
Public Type Archive
FileName As String
Filedata As String
End Type
Dim Pack() As Archive
Public OpendData() As Archive
Public OPendTotal As Long
Public FSize As Long
Public Function CreateArchiv(ByVal Archiv As String) As Long
  Dim F As Integer
  Dim n As Integer
  Dim Filedata As String
  Dim FilePath As String
  Dim FileName As String
  Dim Total As Long
  Dim i As Long
  Dim Head As String
  On Error Resume Next
  If Dir(Archiv) <> "" Then Kill Archiv

 Head = "Pic-Packer1.0 By A.HajaFaizee"
  F = FreeFile
 Total = frmNew.List1.ListCount
 ReDim Preserve Pack(Total)
  Open Archiv For Binary As #F
   
  Put #F, , Total
 
  For i = 0 To Total - 1
 FileName = frmNew.List2.List(i)
 FilePath = frmNew.List1.List(i)
    
 Pack(i).FileName = FileName
 
        n = FreeFile
   Open FilePath For Binary As #n
    Filedata = Space$(LOF(n))
    Get #n, , Filedata
    Close #n
    Pack(i).Filedata = Filedata
    ShowProgress frmNew.Picture1, i, 1, frmNew.List1.ListCount
    DoEvents
  Next i
  Put #F, , Pack
  Put #F, , Head
  Close #F
  

MsgBox "New Archive Created ! And '" & Total & "' Files Put into the Archive", vbInformation, "By A.HajaFaizee"

If Error <> "" Then MsgBox "Error : " & Error
End Function
Public Function OpenArchiv(Archiv As String) As String
   Dim F As Integer
  Dim n As Long
  Dim Filedata As String
  Dim FileName As String
  Dim Total As Long
  Dim i As Long
  Dim Head As String
  Dim R As Boolean
  On Error Resume Next
  Head = Space(29)
  Open Archiv For Binary As 6
  Get 6, LOF(6) - 28, Head
  Close 6
 If Head = "Pic-Packer1.0 By A.HajaFaizee" Then R = True
  If R = True Then
  F = FreeFile
  Open Archiv For Binary As #F
  Get #F, , Total
  Close #F
  ReDim Preserve Pack(Total)
  ReDim Preserve OpendData(Total)
  OPendTotal = Total
  Open Archiv For Binary As #F
  Get #F, , Total
  Get #F, , Pack
  Close #F
  frmMain.List1.Clear
  
  For i = 0 To Total - 1
   FileName = Pack(i).FileName
   n = n + 1
    frmMain.List1.AddItem n & ".  " & FileName
   OpendData(i) = Pack(i)
   
   ShowProgress frmMain.Picture1, n, 1, Total - 1
  Next i
  OpenArchiv = Total
  frmMain.cmdAdd.Enabled = True
  frmMain.cmdDel.Enabled = True
  frmMain.cmdExtract.Enabled = True
  frmMain.Image1.Enabled = True
  Else
  MsgBox Archiv & " Is Not a Valid Packed File"
  End If
  If Error <> "" Then MsgBox "Error in Open an Archive : " & Error
End Function

Public Function ViewArchiv(Index As Long)
Dim PicPath As String
Dim Filedata As String
On Error Resume Next
PicPath = App.Path & "\" & OpendData(Index).FileName
Filedata = OpendData(Index).Filedata
Open PicPath For Binary As 1
  Put 1, , Filedata
  Close 1
  FSize = Len(Filedata)
  frmMain.Image1.Picture = LoadPicture(PicPath)
  Kill PicPath
  If Error <> "" Then MsgBox "Error in Viewing : " & Error
End Function

Public Function Extract(AllF As Boolean, Folder As String) As String
Dim i As Long, n As Long
If Folder <> "\" Then
Folder = Folder & "\"
End If
If AllF = True Then
For i = 0 To OPendTotal - 1
n = n + 1
Open Folder & OpendData(i).FileName For Binary As 1
Put 1, , OpendData(i).Filedata
Close 1
ShowProgress frmMain.Picture1, i, 1, OPendTotal - 1
Next i
Else

For i = 0 To OPendTotal - 1

If frmMain.List1.Selected(i) = True Then
n = n + 1
Open Folder & OpendData(i).FileName For Binary As 1
Put 1, , OpendData(i).Filedata
Close 1
End If
ShowProgress frmMain.Picture1, n, 1, frmMain.List1.SelCount
Next i
End If
Extract = n
End Function
Public Function DeleteF(Totl As Long, Archiv As String) As String
On Error Resume Next
Dim Head As String
Dim i As Long, j As Long
Dim FilN As String
Dim Total As Long
Dim TempData() As Archive
Head = "Pic-Packer1.0 By A.HajaFaizee"
FilN = Right(frmMain.List1.List(i), Len(frmMain.List1.List(i)) - 4)
ReDim Preserve TempData(OPendTotal)
For i = 0 To OPendTotal - 1
If frmMain.List1.Selected(i) = False Then
TempData(j) = OpendData(i)
j = j + 1
End If
Next i

For i = 0 To j
Pack(i) = TempData(i)
Next i
Total = j
Kill Archiv
Open Archiv For Binary As 2
  Put 2, , Total
  Put 2, , Pack
  Put 2, , Head
  Close 2
  If Error <> "" Then MsgBox "Error in Deleting : " & Error
End Function
Public Function AddF(Archiv As String, Totl As Long) As String
Dim Head As String
Dim L As Long
Dim a As Integer
Dim FileName As String
Dim Filedata As String
Dim FilePath As String
Dim Total As Long
Dim TempData() As Archive
Head = "Pic-Packer1.0 By A.HajaFaizee"
 ReDim Preserve TempData(OPendTotal + Totl)
  ReDim Preserve Pack(OPendTotal + Totl)
  a = 0
   For L = 0 To OPendTotal - 1
    TempData(L) = OpendData(L)
    Next L
    
    
    For L = L To OPendTotal + Totl - 1
    FileName = frmNew.List2.List(a)
    FilePath = frmNew.List1.List(a)
    Open FilePath For Binary As 1
    Filedata = String(LOF(1), " ")
    Get 1, , Filedata
    Close 1
   
    TempData(L).Filedata = Filedata
    TempData(L).FileName = FileName
    a = a + 1
    Next L
    
  For L = 0 To OPendTotal + Totl
  Pack(L) = TempData(L)
  Next L
  Total = OPendTotal + Totl
  Open Archiv For Binary As 2
  Put 2, , Total
  Put 2, , Pack
  Put 2, , Head
  Close 2
  AddF = Total
  If Error <> "" Then MsgBox "Error in Adding : " & Error
End Function

Private Sub ShowProgress(picProgress As PictureBox, _
  ByVal Value As Long, _
  ByVal Min As Long, _
  ByVal Max As Long, _
  Optional ByVal bShowProzent As Boolean = True)
  
  Dim pWidth As Long
  Dim intProz As Integer
  Dim strProz As String
  
  ' colors
  Const progBackColor = &HC00000
  Const progForeColor = vbBlack
  Const progForeColorHighlight = vbWhite
  
  ' set Values
  If Value < Min Then Value = Min
  If Value > Max Then Value = Max
  
  ' Prozentwert ausrechnen
  If Max > 0 Then
    intProz = Int(Value / Max * 100 + 0.5)
  Else
    intProz = 100
  End If
    
  With picProgress
    ' check if AutoReadraw=True
    If .AutoRedraw = False Then .AutoRedraw = True
    
    ' clear the picturebox
    picProgress.Cls
    
    If Value > 0 Then
    
      ' calculate barwidth
      pWidth = .ScaleWidth / 100 * intProz
      
      ' Show bar
      picProgress.Line (0, 0)-(pWidth, .ScaleHeight), _
        progBackColor, BF
        
      ' show percent
      If bShowProzent Then
        strProz = CStr(intProz) & " %"
        .CurrentX = (.ScaleWidth - .TextWidth(strProz)) / 2
        .CurrentY = (.ScaleHeight - .TextHeight(strProz)) / 2
      
        ' Foregroundcolor
        If pWidth >= .CurrentX Then
          .ForeColor = progForeColorHighlight
        Else
          .ForeColor = progForeColor
        End If
      
        picProgress.Print strProz
      End If
    End If
  End With
End Sub



