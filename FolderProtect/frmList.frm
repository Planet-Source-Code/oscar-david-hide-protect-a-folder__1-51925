VERSION 5.00
Begin VB.Form frmList 
   BackColor       =   &H00A4AF8F&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1710
      Left            =   150
      TabIndex        =   3
      Top             =   450
      Width           =   7065
   End
   Begin FolderProtect_v2.DMSXpButton cmd 
      Height          =   390
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   2250
      Width           =   1365
      _ExtentX        =   2408
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
      Height          =   390
      Index           =   1
      Left            =   5850
      TabIndex        =   1
      Top             =   2250
      Width           =   1365
      _ExtentX        =   2408
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
      Caption         =   "Close"
      ForeColor       =   12632256
      ForeHover       =   7043413
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select a folder to make it visible:.."
      Height          =   195
      Left            =   75
      TabIndex        =   2
      Top             =   150
      Width           =   2370
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H006B7955&
      BorderWidth     =   3
      Height          =   2790
      Left            =   0
      Top             =   0
      Width           =   7365
   End
End
Attribute VB_Name = "frmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmd_Click(Index As Integer)
    If Index = 1 Then Unload Me Else ShowFolder
End Sub

Private Sub Form_Load()

Dim x As Integer

    Me.Height = Shape1.Height
    Me.Width = Shape1.Width

'    x = SendMessage(List1.hwnd, &H194, 200, ByVal 0&)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    DragForm Button, Me
End Sub

Private Sub Form_Paint()
    ShadeForm Me, &HA4AF8F, 20, 2
End Sub

Private Sub List1_DblClick()
    ShowFolder
End Sub

Sub ShowFolder()

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        ' Set the folder's attribyte to "Normal" or "visible",...
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Dim Fldr As String: Fldr = List1.List(List1.ListIndex)

    If Fldr = "" Then MsgBox "Select a folder!", vbInformation, "huh?": Exit Sub

    SetAttr Fldr, vbArchive     ' Sets the attribute,..

    List1.RemoveItem List1.ListIndex    ' Remove the folder from the list,..

    SaveListBox List1, Log      ' Saves the new Log,..

    MsgBox "The Folder: " & vbCrLf & Fldr & vbCrLf & "  it's visible now,..", vbInformation, "huh?"

End Sub








