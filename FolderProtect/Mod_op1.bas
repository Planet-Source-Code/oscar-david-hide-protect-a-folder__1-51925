Attribute VB_Name = "Mod_op1"
Option Explicit

'*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|
'        for the ini editin',..

    Private Declare Function WritePrivateProfileString Lib "kernel32" Alias _
        "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
        ByVal lpKeyName As Any, ByVal lpString As Any, _
        ByVal lpFileName As String) As Long
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias _
        "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
        ByVal lpKeyName As Any, ByVal lpDefault As String, _
        ByVal lpReturnedString As String, ByVal nSize As Long, _
        ByVal lpFileName As String) As Long
        
'*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|
'       extractin the icons,..
    
    Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" _
        (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As typSHFILEINFO, _
        ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
    Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, _
        ByVal i&, ByVal hDCDest&, ByVal x&, ByVal Y&, ByVal flags&) As Long

'*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|
    
    Private Declare Function SetParent Lib "USER32" (ByVal hWndChild As Long, _
        ByVal hWndNewParent As Long) As Long
    Public Declare Function SendMessage Lib "USER32" Alias "SendMessageA" ( _
        ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
        ByVal lParam As Long) As Long
    Private Declare Function ReleaseCapture Lib "USER32" () As Long
    
    Private Const WM_NCLBUTTONDOWN = &HA1
    Private Const HTCAPTION = 2

'*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|

    Private Declare Function DeleteObject Lib "gdi32" _
       (ByVal hObject As Long) As Long
    Private Declare Function CreateSolidBrush Lib "gdi32" ( _
        ByVal crColor As Long) As Long
    Private Declare Function FillRect Lib "USER32" (ByVal hdc As Long, _
        lpRect As RECT, ByVal hBrush As Long) As Long

'*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|*|


Private Type typSHFILEINFO
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * 260
  szTypeName As String * 80
End Type

Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1
Private Const ILD_TRANSPARENT = &H1
Private Const flags = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Dim FileInfo As typSHFILEINFO

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Public Log As String


Public Sub DragForm(sButton As Integer, Frm As Form)

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        ' Sub that make possible move the form, I got it
            ' from www.areyoufearless.com, I think,..
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    On Local Error Resume Next
    
    If sButton = vbLeftButton Then
    
        Call ReleaseCapture
        Call SendMessage(Frm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)

    End If

End Sub

Public Function ExtractIcon(filename As String, PictureBox As PictureBox, PixelsXY As Integer) As Long
    Dim SmallIcon As Long
    Dim IconIndex As Integer
    If PixelsXY = 16 Then
        SmallIcon = SHGetFileInfo(filename, 0&, FileInfo, Len(FileInfo), flags Or SHGFI_SMALLICON)
    Else
        SmallIcon = SHGetFileInfo(filename, 0&, FileInfo, Len(FileInfo), flags Or SHGFI_LARGEICON)
    End If
    If SmallIcon <> 0 Then
      With PictureBox
        .Height = 15 * PixelsXY
        .Width = 15 * PixelsXY
        .ScaleHeight = 15 * PixelsXY
        .ScaleWidth = 15 * PixelsXY
        .Picture = LoadPicture("")
        .AutoRedraw = True
        SmallIcon = ImageList_Draw(SmallIcon, FileInfo.iIcon, PictureBox.hdc, 0, 0, ILD_TRANSPARENT)
        .Refresh
      End With
    End If
End Function

Public Function WRITEINI(ByVal SECTION As String, KEY As String, value As String, INIFILE As String)
WritePrivateProfileString SECTION, KEY, value, INIFILE
End Function
Public Function GETINI(ByVal SECTION As String, KEY As String, INIFILE As String, ByRef rVALUE As String) As Boolean
Dim value As String * 256
Dim a As Long
a = GetPrivateProfileString(SECTION, _
    KEY, "?!?", value, 256, _
    INIFILE)
If Left$(value, 3) = "?!?" Then
    GETINI = False
Else
    GETINI = True
    rVALUE = Left$(value, a)
End If
End Function



'
' Shades the form in a similar manner to many
' install programs.
'
' Optional Arguments:
' StartColor is what color to start with.
'   (Default = vbBlue)
' Fstep is the number of steps to use to fill the form.
'   (Default = 64)
' Cstep is the color step (change in color per step).
'   (Default = 4)
'
' Note: the effect can be reversed by calling ShadeForm with
'    a StartColor near black (but not completely 0) and by
'    setting a negative color step.
'
Public Sub ShadeForm(F As Form, Optional StartColor As Variant, Optional Fstep As Variant, Optional Cstep As Variant)
   Dim FillStep As Single  ' Not an integer because sometimes
                           ' rounding leaves a large bottom region
   
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            ' I dont' remember where I got this _
                sub, I think from www.thevbzone.com
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

   
   Dim c As Long
   Dim FillArea As RECT
   Dim i As Integer
   Dim oldm As Integer
   Dim hBrush As Long
   Dim c2(1 To 3) As Long
   Dim cs2(1 To 3) As Long
   Dim fs As Long
   Dim cs As Integer
      
   ' Set defaults
   fs = IIf(IsMissing(Fstep), 64, CLng(Fstep))
   cs = IIf(IsMissing(Cstep), 4, CInt(Cstep))
   c = IIf(IsMissing(StartColor), vbBlue, CLng(StartColor))
   
   
   oldm = F.ScaleMode
   F.ScaleMode = vbPixels
   FillStep = F.ScaleHeight / fs
   FillArea.Left = 0
   FillArea.Right = F.ScaleWidth
   FillArea.Top = 0

   ' Break down the color and set individual
   ' color steps
   c2(1) = c And 255#
   cs2(1) = IIf(c2(1) > 0, cs, 0)
   c2(2) = (c \ 256#) And 255#
   cs2(2) = IIf(c2(2) > 0, cs, 0)
   c2(3) = (c \ 65536#) And 255#
   cs2(3) = IIf(c2(3) > 0, cs, 0)
   
   
   For i = 1 To fs
      FillArea.Bottom = FillStep * i

      hBrush = CreateSolidBrush(RGB(c2(1), c2(2), c2(3)))
      FillRect F.hdc, FillArea, hBrush
      DeleteObject hBrush
      
      ' Could do this in a loop, but it's simple
      ' and may be faster.
      c2(1) = (c2(1) - cs2(1)) And 255#
      c2(2) = (c2(2) - cs2(2)) And 255#
      c2(3) = (c2(3) - cs2(3)) And 255#
      
      FillArea.Top = FillArea.Bottom
   Next i
   
   F.ScaleMode = oldm
   
   
End Sub

Public Sub SaveListBox(TheList As ListBox, Directory As String)

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            ' The save & load  listboxes  sub's _
                are form : Markku Strömberg
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
    Dim SaveList As Long
    On Error Resume Next
    Open Directory$ For Output As #1


    For SaveList& = 0 To TheList.ListCount - 1
        Print #1, TheList.List(SaveList&)
    Next SaveList&
    Close #1
End Sub

Public Sub LoadListBox(TheList As ListBox, Directory As String)
    Dim MyString As String
    On Error Resume Next
    
    Open Directory$ For Input As #1
    
    While Not EOF(1)
'Do Until EOF(1)
    
    
    ' I changed "Input" for "Line Input" since
        ' the sub does not load long paths correctly,..
    
         Line Input #1, MyString$

        DoEvents
        
            TheList.AddItem MyString$
            
'Loop
        Wend
        
        Close #1
        
    End Sub
