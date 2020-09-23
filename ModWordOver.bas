Attribute VB_Name = "ModWordOver"
Option Explicit
'Simple Word Over function for Textbox not Richtext
'© 2005 Scythe
'www.scythe-tools.de


Private Type POINTAPI
 X As Long
 Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long

Private Const EM_CHARFROMPOS = &HD7

'Find the Word the Cursor is over in a Textbox
Public Function TextWordOver(ByRef TextBox As TextBox) As String
 Dim PT      As POINTAPI
 Dim CurLine As Long
 Dim CurPosX As Long
 Dim StartX  As Long
 Dim EndX    As Long
 Dim Tmp     As String

With TextBox
 GetCursorPos PT
 ScreenToClient .hwnd, PT

 CurLine = SendMessage(.hwnd, EM_CHARFROMPOS, 0, ByVal CLng(CInt(PT.Y) * &H10000 + CInt(PT.X)))

 CurPosX = CurLine - (CurLine And &HFFFF0000) + 1
 'This will give you the Linenumber (I dont need it)
 'CurLine = (lCurLine \ &H10000)

 'Search for the Start
 For StartX = CurPosX To 1 Step -1
  Tmp = Mid$(TextBox.Text, StartX, 1)
  'Check for Words
  If Not ((Tmp >= "-" And Tmp <= ":") Or (Tmp >= "a" And Tmp <= "z") Or (Tmp >= "A" And Tmp <= "Z") Or Tmp = "_") Then Exit For
  'If you only want words not webadresses
  'If Not ((Tmp >= "0" And Tmp <= "9") Or (Tmp >= "a" And Tmp <= "z") Or (Tmp >= "A" And Tmp <= "Z")) Then Exit For
 Next StartX
 StartX = StartX + 1

 'Search the end
 For EndX = CurPosX To Len(TextBox.Text)
  Tmp = Mid$(TextBox.Text, EndX, 1)
  'Check for Words
  If Not ((Tmp >= "-" And Tmp <= "9") Or (Tmp >= "a" And Tmp <= "z") Or (Tmp >= "A" And Tmp <= "Z") Or Tmp = "_") Then Exit For
 Next EndX
 EndX = EndX - 1

 'Set the word we found
 If StartX <= EndX Then TextWordOver = Mid$(TextBox.Text, StartX, EndX - StartX + 1)
End With
End Function

