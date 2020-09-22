Attribute VB_Name = "IniModule"
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Declare Function ShellExecute _
   Lib "shell32.dll" Alias "ShellExecuteA" _
   (ByVal hwnd As Long, _
   ByVal lpOperation As String, _
   ByVal lpFile As String, _
   ByVal lpParameters As String, _
   ByVal lpDirectory As String, _
   ByVal nShowCmd As Long) As Long

Public Const SW_SHOWNORMAL = 1
Public Const SWW_HPARENT = -8

Public Const LB_FINDSTRING = &H18F
Public Const EM_UNDO = &HC7
