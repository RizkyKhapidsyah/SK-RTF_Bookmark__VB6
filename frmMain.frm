VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmMain 
   Caption         =   "RTF Bookmark Example"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin vbalIml6.vbalImageList ilsIcons 
      Left            =   1800
      Top             =   960
      _ExtentX        =   953
      _ExtentY        =   953
      ColourDepth     =   4
      Size            =   4700
      Images          =   "frmMain.frx":0000
      KeyCount        =   5
      Keys            =   "SPACERÿDELETEÿPREVIOUSÿNEXTÿTOGGLE"
   End
   Begin RichTextLib.RichTextBox rtfText 
      Height          =   735
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmMain.frx":127C
   End
   Begin VB.Menu mnuBookmarkTOP 
      Caption         =   "&Bookmarks"
      Begin VB.Menu mnuBookmarks 
         Caption         =   "&Toggle Bookmark"
         Index           =   0
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuBookmarks 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuBookmarks 
         Caption         =   "&Next Bookmark"
         Enabled         =   0   'False
         Index           =   2
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuBookmarks 
         Caption         =   "&Previous Bookmark"
         Enabled         =   0   'False
         Index           =   3
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuBookmarks 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuBookmarks 
         Caption         =   "&Delete All Bookmarks"
         Enabled         =   0   'False
         Index           =   5
         Shortcut        =   ^D
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'============================================================================================
' RTF Bookmark Example
'============================================================================================
' Created By: Marc Cramer
' Published Date: 12/22/2000
' WebSite: MKC Computers at http://www.mkccomputers.com
'============================================================================================
' Additional Controls: vbalImageList
' Additional Reference: vbAccelerator IconMenu.dll
' WebSite Downloaded From: VBAccelerator at http://www.vbaccelerator.com
'============================================================================================
' NOTES:
' do not close the program while running from the IDE...VB will crash because of menu icons
'============================================================================================
Option Explicit

Dim MyMenu As cIconMenu 'menus with icons
Dim MyBookmarks As clsArray 'bookmark array
Const DebugMode = False 'used for debug only
'===============================================================================
Private Sub Form_Load()
' load form and create instances of classes and add images to menu
  Set MyBookmarks = New clsArray
  Set MyMenu = New cIconMenu
  ' this following adds the images to the menus
  With MyMenu
    .Attach Me.hWnd
    .ImageList = ilsIcons
    .IconIndex("&Delete All Bookmarks") = ilsIcons.ItemIndex("DELETE")
    .IconIndex("&Previous Bookmark") = ilsIcons.ItemIndex("PREVIOUS")
    .IconIndex("&Next Bookmark") = ilsIcons.ItemIndex("NEXT")
    .IconIndex("&Toggle Bookmark") = ilsIcons.ItemIndex("TOGGLE")
  End With
  ' load something into rtfText for example purposes
  rtfText.LoadFile App.Path & "\temp.txt"
  '=====================================================
  ' DEBUG MODE
  '=====================================================
    If DebugMode = True Then DebugMessage ("Menu Icons Created")
  '=====================================================
End Sub 'Form_Load()
'===============================================================================
Private Sub Form_Resize()
On Error Resume Next
' resize the rtfbox to proper size
  rtfText.Move 25, 25, Me.ScaleWidth - 50, Me.ScaleHeight - 50
  rtfText.RightMargin = rtfText.Width - 125
End Sub 'Form_Resize()
'===============================================================================
Private Sub mnuBookmarks_Click(Index As Integer)
' add, delete, move next, move previous based on user selection
  Dim Counter As Integer
  Dim CurrentLine As Integer
  Dim CurrentPosition As Integer
  
  CurrentLine = CurrentLineStartPos(rtfText)
  CurrentPosition = rtfText.SelStart
  With MyBookmarks
    Select Case Index
      Case 0: 'toggle bookmark
        Dim ItemIndex As Integer
        ItemIndex = .FindItemIndex(CurrentLine)
        If ItemIndex = -1 Then
          .AddNew CurrentLine
          rtfText.SelStart = CurrentLineStartPos(rtfText)
          rtfText.SelLength = CurrentLineLength(rtfText)
          rtfText.SelColor = vbRed
          rtfText.SelStart = CurrentPosition
        '=====================================================
        ' DEBUG MODE
        '=====================================================
          If DebugMode = True Then DebugMessage ("ValueAdded: " & CurrentLine)
        '=====================================================
        Else
          .Delete ItemIndex
          rtfText.SelStart = CurrentLineStartPos(rtfText)
          rtfText.SelLength = CurrentLineLength(rtfText)
          rtfText.SelColor = vbBlack
          rtfText.SelStart = CurrentPosition
        '=====================================================
        ' DEBUG MODE
        '=====================================================
          If DebugMode = True Then DebugMessage ("Deleted Index: " & ItemIndex)
        '=====================================================
        End If
        If .IndexCount >= 0 Then .Sort_BubbleSort 'only need to sort if more then one bookmark
        rtfText.SetFocus
      Case 2: 'next bookmark
        Counter = 0
        If CurrentLine < .Value(.IndexCount) Then
          Do Until .Value(Counter) > rtfText.SelStart
            Counter = Counter + 1
          Loop
        End If
        rtfText.SelStart = .Value(Counter)
        rtfText.SetFocus
        '=====================================================
        ' DEBUG MODE
        '=====================================================
          If DebugMode = True Then DebugMessage ("Moved Next")
        '=====================================================
      Case 3: 'previous bookmark
        Counter = .IndexCount
        If CurrentLine > .Value(0) Then
          Do Until .Value(Counter) < CurrentLineStartPos(rtfText)
            Counter = Counter - 1
          Loop
        End If
        rtfText.SelStart = .Value(Counter)
        rtfText.SetFocus
        '=====================================================
        ' DEBUG MODE
        '=====================================================
          If DebugMode = True Then DebugMessage ("Moved Previous")
        '=====================================================
      Case 5: 'clear all bookmarks
        .DeleteAll
        rtfText.SelStart = 0
        rtfText.SelLength = Len(rtfText.Text)
        rtfText.SelColor = vbBlack
        rtfText.SelStart = CurrentPosition
        rtfText.SetFocus
        '=====================================================
        ' DEBUG MODE
        '=====================================================
          If DebugMode = True Then DebugMessage ("Deleted All")
        '=====================================================
      Case Else:
        'not a valid choice so do nothing
    End Select
  End With
  rtfText.SelColor = vbBlack 'just to make sure that as we type it is in correct color
  UpdateMenu
End Sub 'mnuBookmarks_Click(Index As Integer)
'===============================================================================
Private Sub rtfText_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
' right click popup menu
  If Button = vbRightButton Then PopupMenu mnuBookmarkTOP
End Sub 'rtfText_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'===============================================================================
Private Sub UpdateMenu()
' enable or disable menu choices based on array size
  If MyBookmarks.IndexCount >= 0 Then
    mnuBookmarks(2).Enabled = True
    mnuBookmarks(3).Enabled = True
    mnuBookmarks(5).Enabled = True
    '=====================================================
    ' DEBUG MODE
    '=====================================================
      If DebugMode = True Then DebugMessage ("Menu Enabled")
    '=====================================================
  Else
    mnuBookmarks(2).Enabled = False
    mnuBookmarks(3).Enabled = False
    mnuBookmarks(5).Enabled = False
    '=====================================================
    ' DEBUG MODE
    '=====================================================
      If DebugMode = True Then DebugMessage ("Menu Disabled")
    '=====================================================
  End If
End Sub 'UpdateMenu()
'===============================================================================
Private Sub DebugMessage(MyMessage As String)
' print debug message if Const DebugMode = True
  Debug.Print MyMessage
End Sub 'DebugMessage(MyMessage As String)
'===============================================================================

