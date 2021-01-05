Attribute VB_Name = "modRTF"
'============================================================================================
' SendMessage API
'============================================================================================
' Adapted and Modified By: Marc Cramer
' Published Date: 12/22/2000
' WebSite: MKC Computers at http://www.mkccomputers.com
'============================================================================================
' Based On: Extending the Textbox Control Demonstration Project - By Joacim Andersson
' Published Date: 8/5/99
' WebSite: VB-World.net at http://www.vb-world.net
'=========================================================================================
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias _
        "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, lParam As Any) As Long

Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_GETLINE = &HC4
Private Const EM_LINELENGTH = &HC1
Private Const EM_LINEINDEX = &HBB

Public Property Get CurrentLinePos(rtf As RichTextBox) As Long
    CurrentLinePos = SendMessage(rtf.hWnd, EM_LINEFROMCHAR, ByVal -1, 0&)
End Property

Public Property Get CurrentLineLength(rtf As RichTextBox) As Long
    CurrentLineLength = SendMessage(rtf.hWnd, EM_LINELENGTH, rtf.SelStart, 0)
End Property

Public Property Get CurrentLineStartPos(rtf As RichTextBox) As Long
    CurrentLineStartPos = SendMessage(rtf.hWnd, EM_LINEINDEX, -1, 0&)
End Property


