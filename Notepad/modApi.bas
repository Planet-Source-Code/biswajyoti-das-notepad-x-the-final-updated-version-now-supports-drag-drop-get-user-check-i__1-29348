Attribute VB_Name = "modApi"
' Win32 Declarations for Cut, Copy, Paste and Delete
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_CUT = &H300
Public Const WM_COPY = &H301
Public Const WM_PASTE = &H302
Public Const WM_CLEAR = &H303
Public Const WM_USER = &H400
Public Const EM_CANUNDO = &HC6
Public Const EM_UNDO = &HC7

Public Const EM_LINEINDEX = &HBB
Private Const EM_SETTARGETDEVICE = (WM_USER + 72)
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_LINEFROMCHAR = &HC9
