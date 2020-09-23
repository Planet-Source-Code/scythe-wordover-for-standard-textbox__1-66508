VERSION 5.00
Begin VB.Form FrmWordOverTest 
   Caption         =   "Word Over Test"
   ClientHeight    =   11610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   ScaleHeight     =   11610
   ScaleWidth      =   11475
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text1 
      Height          =   11295
      Left            =   0
      Locked          =   -1  'True
      MouseIcon       =   "FrmWordOverTest.frx":0000
      MousePointer    =   1  'Pfeil
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "FrmWordOverTest.frx":0152
      Top             =   0
      Width           =   11415
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   11280
      Width           =   11535
   End
End
Attribute VB_Name = "FrmWordOverTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Simple WordOver for Textbox not for Richtext
'© 2005 Scythe

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1



Private Sub Text1_Click()
 Dim WordOver As String

 'Get the word under the cursor
 WordOver = TextWordOver(Text1)

 'open a webpage
 If InStr(1, LCase(WordOver), "www.") <> 0 Or InStr(1, LCase(WordOver), "http://") <> 0 Or InStr(1, LCase(WordOver), "ftp://") <> 0 Then
  ShellExecute Me.hwnd, "open", WordOver, vbNullString, vbNullString, SW_SHOWNORMAL
 End If

End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim WordOver As String

 'Get the word under the cursor
 WordOver = TextWordOver(Text1)

 'Show link cursor
 If InStr(1, LCase(WordOver), "www.") <> 0 Or InStr(1, LCase(WordOver), "http://") <> 0 Or InStr(1, LCase(WordOver), "ftp://") <> 0 Then
  Text1.MousePointer = 99
 Else
  Text1.MousePointer = 0
 End If

 'Show the Word
 Label1 = WordOver
End Sub
