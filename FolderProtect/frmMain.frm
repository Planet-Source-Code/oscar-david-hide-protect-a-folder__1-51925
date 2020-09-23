VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00A4AF8F&
   BorderStyle     =   0  'None
   Caption         =   "FolderProtect v2.0"
   ClientHeight    =   2130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4815
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1200
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   16
      Top             =   4650
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   375
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   15
      Top             =   4650
      Width           =   495
   End
   Begin VB.ComboBox cmb1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   225
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2475
      Width           =   1440
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   225
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Select a folder ----->"
      Top             =   1425
      Width           =   2715
   End
   Begin FolderProtect_v2.DMSXpButton cmd 
      Height          =   390
      Index           =   1
      Left            =   1950
      TabIndex        =   1
      Top             =   2175
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Lock Folder"
      ForeColor       =   12632256
      ForeHover       =   7043413
   End
   Begin FolderProtect_v2.DMSXpButton cmd 
      Height          =   390
      Index           =   2
      Left            =   1950
      TabIndex        =   2
      Top             =   2625
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Unlock Folder"
      ForeColor       =   12632256
      ForeHover       =   7043413
   End
   Begin FolderProtect_v2.DMSXpButton cmd 
      Height          =   315
      Index           =   0
      Left            =   3075
      TabIndex        =   0
      Top             =   1425
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "..."
      ForeColor       =   0
      ForeHover       =   9280371
   End
   Begin FolderProtect_v2.DMSXpButton cmd 
      Height          =   390
      Index           =   3
      Left            =   225
      TabIndex        =   5
      Top             =   3600
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Hide Folder"
      ForeColor       =   12632256
      ForeHover       =   7043413
   End
   Begin FolderProtect_v2.DMSXpButton cmd 
      Height          =   390
      Index           =   4
      Left            =   1950
      TabIndex        =   6
      Top             =   3600
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Show Folder"
      ForeColor       =   12632256
      ForeHover       =   7043413
   End
   Begin FolderProtect_v2.DMSXpButton cmd 
      Height          =   315
      Index           =   5
      Left            =   2175
      TabIndex        =   10
      Top             =   4350
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Open icon"
      ForeColor       =   12632256
      ForeHover       =   7043413
   End
   Begin FolderProtect_v2.DMSXpButton cmd 
      Height          =   390
      Index           =   6
      Left            =   1950
      TabIndex        =   11
      Top             =   4800
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Change Icon"
      ForeColor       =   12632256
      ForeHover       =   7043413
   End
   Begin FolderProtect_v2.DMSXpButton cmd 
      Height          =   390
      Index           =   7
      Left            =   1950
      TabIndex        =   12
      Top             =   5250
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Set Deafult Icon"
      ForeColor       =   12632256
      ForeHover       =   7043413
   End
   Begin FolderProtect_v2.DMSXpButton cmd 
      Height          =   390
      Index           =   8
      Left            =   225
      TabIndex        =   13
      Top             =   6000
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "About,.."
      ForeColor       =   7043413
      ForeHover       =   10792847
   End
   Begin FolderProtect_v2.DMSXpButton cmd 
      Height          =   390
      Index           =   9
      Left            =   1950
      TabIndex        =   14
      Top             =   6000
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Exit"
      ForeColor       =   7043413
      ForeHover       =   8421631
   End
   Begin VB.Image Image2 
      Height          =   450
      Left            =   3000
      Picture         =   "frmMain.frx":0ECA
      Stretch         =   -1  'True
      Top             =   75
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H006B7955&
      BorderWidth     =   2
      Index           =   4
      X1              =   0
      X2              =   3675
      Y1              =   1275
      Y2              =   1275
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H006B7955&
      BorderWidth     =   3
      Height          =   6540
      Left            =   0
      Top             =   0
      Width           =   3690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Icon"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   4
      Left            =   1275
      TabIndex        =   18
      Top             =   5175
      Width           =   420
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Icon"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   3
      Left            =   375
      TabIndex        =   17
      Top             =   5175
      Width           =   525
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line1 
      BorderColor     =   &H006B7955&
      BorderWidth     =   2
      Index           =   3
      X1              =   150
      X2              =   3525
      Y1              =   5850
      Y2              =   5850
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change folder icon,.."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   375
      TabIndex        =   9
      Top             =   4275
      Width           =   1695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H006B7955&
      BorderWidth     =   2
      Index           =   2
      X1              =   150
      X2              =   3525
      Y1              =   4125
      Y2              =   4125
   End
   Begin VB.Line Line1 
      BorderColor     =   &H006B7955&
      BorderWidth     =   2
      Index           =   1
      X1              =   150
      X2              =   3525
      Y1              =   3150
      Y2              =   3150
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Simple hide foder,.."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   450
      TabIndex        =   8
      Top             =   3300
      Width           =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Simple lock foder,.."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   225
      TabIndex        =   7
      Top             =   2025
      Width           =   1560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H006B7955&
      BorderWidth     =   2
      Index           =   0
      X1              =   150
      X2              =   3525
      Y1              =   1875
      Y2              =   1875
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   0
      Picture         =   "frmMain.frx":1D94
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3750
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oIcon As String
Dim Folder As String

Private Sub Form_Load()

    Me.Height = Shape1.Height
    Me.Width = Shape1.Width
    
' Fill the combo box,..
    cmb1.List(0) = "Word Document"
    cmb1.List(1) = "XML Document"
    cmb1.List(2) = "Recycled Bin"
    cmb1.List(3) = "Control Panel"
    cmb1.List(4) = "Printers"
    cmb1.List(5) = "Audio File"
    cmb1.List(6) = "Video File"
    cmb1.List(7) = "Network"
    
    Text1.Text = App.path: Folder = Text1.Text
    
' Initialize the log file,..
    Log = Slash(App.path) & "l0g.ext"

    ExtractIcon Folder, Picture1, 32

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    DragForm Button, Me     ' to move the form,..
    
End Sub

Private Sub cmd_Click(Index As Integer)

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        ' Here we have all the "commands", just _
            follow the sub's,...
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Select Case Index
    
    Case 0
        FolderIcon
    Case 1
        LockFolder
    Case 2
        UnlockFolder
    Case 3
        HideFolder
    Case 4
        ShowList
    Case 5
        LoadIcon
    Case 6
        ChangeIcon
    Case 7
        SetDefault
    Case 8
            MsgBox "    Simple Hide/Protect" _
            & vbCrLf & "         Folder, made " _
            & vbCrLf & "            in VB by:" _
            & vbCrLf _
            & vbCrLf & "          [ scodman ]" _
            & vbCrLf _
            , , App.Title
        
    Case 9
        Unload Me
        End

End Select

End Sub
Sub UnlockFolder()

Dim NewFolder As String: NewFolder = GetFilePath(Folder) & "\" & NoExtPath(Folder)  ' Gets the path & the name of the carpet,..

    Name Folder As NewFolder    ' Renames the folder to the normal one,..

    MsgBox "Folder: " & vbCrLf & Folder & vbCrLf & "Unlocked!", vbInformation

   Text1.Text = NewFolder

   ExtractIcon NewFolder, Picture1, 32     ' reload the icon of the folder,..
   
   FillVar  ' Fill the "Folder" variable,..

End Sub

Sub LockFolder()

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        ' Renames the Folder to "FOLDER.{....."
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
If cmb1.ListIndex = -1 Then Beep: Exit Sub

Dim NewFolder As String

Dim TypeOfFolder(0 To 7) As String, X As Integer

    TypeOfFolder(0) = ".{21EC2020-3AEA-1069-A2DD-08002B30309D}"     'Word Document
    TypeOfFolder(1) = ".{2227A280-3AEA-1069-A2DE-08002B30309D}"     'XML Document
    TypeOfFolder(2) = ".{0003000D-0000-0000-C000-000000000046}"     'Recycled Bin"
    TypeOfFolder(3) = ".{00022602-0000-0000-C000-000000000046}"     'Control Panel
    TypeOfFolder(4) = ".{9E56BE61-C50F-11CF-9A2C-00A0C90A90CE}"    'Printers
    TypeOfFolder(5) = ".{00020906-0000-0000-C000-000000000046}"     'Audio File
    TypeOfFolder(6) = ".{48123bc4-99d9-11d1-a6b3-00c04fd91555}"     'Video File
    TypeOfFolder(7) = ".{645FF040-5081-101B-9F08-00AA002F954E}"      'Network

    X = cmb1.ListIndex
    
    NewFolder = Folder & TypeOfFolder(X)    ' Accordin to the type selected,..
    
        Name Folder As NewFolder    ' renames the folder.
    
        MsgBox "Folder: " & vbCrLf & Folder & vbCrLf & "Locked!", vbInformation

        Text1.Text = NewFolder
        
        ExtractIcon NewFolder, Picture1, 32
        
        FillVar     ' again, fill the var,..
        
End Sub

Sub HideFolder()

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        ' The only thing we do here is set the attributes
                'to the folder to System & Hidden,....
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    SetAttr Folder, vbHidden + vbSystem     ' Set the attributes,..

        Open Log For Append As #1
            Print #1, Folder        ' Saves the folder hidden to the log,..
        Close #1
    
    MsgBox "The folder: " & vbCrLf & Folder & vbCrLf & "  it's not visible now", vbInformation, "yeah!"

FillVar     ' yeah,.....

End Sub

Sub ShowList()

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        ' Load the folder log to the list form & show it
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Dim X As Integer

    If Exist(Log) Then  ' If the log does exist,..
    
        LoadListBox frmList.List1, Log    ' Load the log
    
'        x = SendMessage(frmList.List1.hwnd, &H194, 200, ByVal 0&)
    
        frmList.Show vbModal, Me    ' huh?,..

    End If

End Sub


Sub FolderIcon()

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        ' Display the folder's icon selected to the pic box,..
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

On Error Resume Next

'    Dim Fldr As String
    
        Folder = BrowseForFolder(Me.hwnd, "Select a folder to edit")    ' Opens the Select A Folder dialog,..
        
        If Len(Folder) > 3 Then     ' if the folder it's not a drive,..
            
            Text1.Text = Folder
            
            ExtractIcon Folder, Picture1, 32    ' extracs the folder's icon,..

        End If
        
End Sub

Sub SetDefault()

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        ' Sets the defult icon of the folder,..
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~

On Error Resume Next

FillVar

    Kill Folder & "\desktop.ini"    ' Deletes the "INI" file so the "link" to the icon no longer exist,..
    
    SetAttr Folder, vbArchive
    
    ExtractIcon Folder, Picture1, 32


End Sub


Sub LoadIcon()

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        ' Open a selected icon to manage with
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    oIcon = ShowOpen(Me.hwnd, "Icon files,Exe files" _
                + vbNullChar + "*.ico;*.exe", "Select an Icon,..")

    If Exist(oIcon) Then ExtractIcon oIcon, Picture2, 32

End Sub

Function Exist(path As String) As Boolean

    '~~~~~~~~~~~~~~~~~~~~~~~~
        ' Check for the existing of a file,..
    '~~~~~~~~~~~~~~~~~~~~~~~~

On Error GoTo w

    FileLen (path): Exist = True: Exit Function
    
w:     Exist = False

End Function

Function NoExtPath(ByVal strSource As String) As String

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        ' Returns the file name without path & extension,..
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    On Error GoTo Error_Handler
        
    Dim lngBegin As Long
    Dim lngToRight As Long
    
        lngBegin = InStr(1, StrReverse(strSource), "\", vbTextCompare)
        lngToRight = InStr(1, StrReverse(strSource), ".", vbTextCompare)
        
    NoExtPath = Left(Right(strSource, lngBegin - 1), lngBegin - lngToRight - 1)
    
    Exit Function
    
Error_Handler:

    NoExtPath = ""
    
End Function

Private Sub ChangeIcon()

    If Len(Folder) > 3 Then  'Check selected is not a drive
    
        WRITEINI ".ShellClassInfo", "iconfile", oIcon, Folder & "\desktop.ini"
        WRITEINI ".ShellClassInfo", "iconindex", "0", Folder & "\desktop.ini"
        
        SetAttr Folder, vbSystem 'This is important
        SetAttr Folder & "\desktop.ini", vbHidden + vbSystem
        
        MsgBox "Icon Changed", vbInformation, App.Title
        ExtractIcon Folder, Picture1, 32
    
    End If
    
    FillVar
    
End Sub

Public Function GetFilePath(filename As String)
    Dim X As String
    Dim position As Integer
    Dim Y As Integer
    Dim myPath As String
    X = filename
    For Y = Len(X) To 1 Step -1
        If Mid(X, Y, 1) = "\" Then
            position = Y - 1
            Exit For
        End If
    Next Y
myPath = Left(X, position)
GetFilePath = Format(myPath)
End Function


Private Sub Form_Paint()
'   Drawin the gradient form,..
    ShadeForm Me, &HA4AF8F, 20, 2
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragForm Button, Me
End Sub

Function Slash(strPath As String) As String
'Receives the App.Path and returns a path containing backslash
    If Right$(strPath, 1) <> "\" Then Slash = strPath & "\" Else Slash = strPath
End Function

Sub FillVar()
    Folder = Text1.Text
End Sub



