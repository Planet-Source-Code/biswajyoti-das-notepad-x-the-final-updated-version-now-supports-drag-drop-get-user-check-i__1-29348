VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEdit 
   Caption         =   "Untitled - Notepadx"
   ClientHeight    =   7845
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11190
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Notepad"
   LockControls    =   -1  'True
   ScaleHeight     =   7845
   ScaleWidth      =   11190
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEdit 
      DragIcon        =   "Form1.frx":014A
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Index           =   1
      Left            =   0
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtEdit 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Index           =   0
      Left            =   0
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   2295
   End
   Begin MSComDlg.CommonDialog dlgEdit 
      Left            =   3960
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   5
   End
   Begin VB.Menu mnu0 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnu5 
         Caption         =   "-"
      End
      Begin VB.Menu mnu6 
         Caption         =   "Page Se&tup"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnu8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "Undo"
         Enabled         =   0   'False
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnu21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "De&lete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnu16 
         Caption         =   "-"
      End
      Begin VB.Menu mnu17 
         Caption         =   "Select &All"
      End
      Begin VB.Menu mnu18 
         Caption         =   "Time/&Date"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnu19 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditWordWrap 
         Caption         =   "&Word Wrap"
      End
      Begin VB.Menu mnuEditSelectFont 
         Caption         =   "Set &Font..."
      End
   End
   Begin VB.Menu mnu22 
      Caption         =   "&Search"
      Begin VB.Menu mnu23 
         Caption         =   "&Find..."
      End
      Begin VB.Menu mnu24 
         Caption         =   "Find &Next"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnu25 
      Caption         =   "&Help"
      Begin VB.Menu mnu26 
         Caption         =   "&About  Notepadx"
      End
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Please vote for this code .. I have spent almost 5 hours on it
'You can feel the difference when you see all the WordWrap and other feature
'that most other avoid ... but I just wanted it to be a clone so I didn't
'include any advance text searching .. I have done simple one
'Last but not least I wanna thank John for telling me to do the .LOG thing of
'notepad .. so what is .LOG .. just save any file and include a line writing .LOG at
'the first line of your file and when you reopen the file you will see the current
'date and time right at the bottom of your file .... just remember since the textbox
'can hold only 65536 bytes so if your file exceeds that size when you add the .LOG thing
'at the top of your file ... the date/time will not be shown at the end .. thsi is not
'bug .. I suppose :-(. I have also corrected the font bug and also the word wrap bug ...
'when you selected word wrap the font of the text box was not changing now it does ....
'Please report any bugs and please please vote for this code .. I badly need it
'Thanks for downloading my code ..
'Biswajyoti Das (askbiswa@hotmail.com)

Option Explicit
Dim pos As Integer
Dim place As Integer
Dim MOST_IMP As Integer
Dim ANOTHER_ONE As Integer
Dim strImp As String
Dim strTitle As String
Dim blnChange As Boolean, blnCancelSave As Boolean
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, _
    ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Sub Form_Load()
dlgEdit.FontName = txtEdit(0).Font
pos = 1
place = 0
  
If txtEdit(0).Visible = True Then
 MOST_IMP = 0
Else
 MOST_IMP = 1
End If
txtEdit(MOST_IMP).Height = Me.ScaleHeight
txtEdit(MOST_IMP).Width = Me.ScaleWidth
End Sub
Private Sub Form_Resize()
Call Form_Load
End Sub
Private Sub mnu17_Click()
txtEdit(MOST_IMP).SelStart = 0
txtEdit(MOST_IMP).SelLength = 65535   'max length
End Sub

Private Sub mnu18_Click()
txtEdit(MOST_IMP).Text = txtEdit(MOST_IMP).Text & " " & VBA.time & " " & Date
Call keybd_event(vbKeyEnd, 1, 0, 0) 'I hope you understand why I used this :-)
End Sub

Private Sub mnu20_Click()
dlgEdit.ShowFont
End Sub
Private Sub mnu23_Click()
 ShowFind Me, FR_DOWN, ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Const conBtns As Integer = vbYesNoCancel + vbExclamation _
                            + vbDefaultButton3 + vbApplicationModal
    Dim conMsg As String
    Dim intUserResponse As Integer
    Me.GetText
    conMsg = "The text in the " & strTitle & " file has changed." & vbCrLf & vbCrLf & "Do you want to save the changes ?"
    If blnChange = True Then                'document was changed since last save
        intUserResponse = MsgBox(conMsg, conBtns, "Notepadx")
        Select Case intUserResponse
            Case vbYes                      'user wants to save current document
                Call mnuFileSave_Click
                If blnCancelSave = True Then    'user canceled save
                    Cancel = 1              'return to document-don't unload form
                End If
            Case vbNo                       'user does not want to save current document
                'unload form and exit
            Case vbCancel
                Cancel = 1                  'return to document-don't unload form
        End Select
    End If
    
End Sub

Private Sub mnu24_Click()
strSearch (strImp)
End Sub

Private Sub mnu26_Click()
Form2.Show vbModal
End Sub

Private Sub mnuEditDelete_Click()
 txtEdit(MOST_IMP).SelText = ""
End Sub

Private Sub mnuEditSelectFont_Click()
    ' Set flags to show both printer and screen fonts
    ' Alternately, use the flag cdlCFPrinterFonts or cdlCFScreenFonts to show a specific set
    dlgEdit.flags = cdlCFBoth
    dlgEdit.ShowFont
    ' Display selected font
    With txtEdit(MOST_IMP)
     .FontName = dlgEdit.FontName
     .FontSize = dlgEdit.FontSize
    End With
End Sub

Private Sub mnuEditUndo_Click()
SendMessage txtEdit(MOST_IMP).hwnd, EM_UNDO, 0, 0&
End Sub

Private Sub mnuEditWordWrap_Click()

mnuEditWordWrap.Checked = Not mnuEditWordWrap.Checked
If mnuEditWordWrap.Checked = True Then
 ANOTHER_ONE = 1
 MOST_IMP = 1
 blnChange = False
 mnuEditUndo.Enabled = False
 txtEdit(0).Visible = False
 txtEdit(1).Visible = True
 txtEdit(1).Font = dlgEdit.FontName
 txtEdit(1).FontSize = dlgEdit.FontSize
 txtEdit(1).Height = Me.ScaleHeight - 10
 txtEdit(1).Width = Me.ScaleWidth - 10
 txtEdit(1) = txtEdit(0)
 txtEdit(0) = ""
Else
 MOST_IMP = 0
 ANOTHER_ONE = 1
 blnChange = False
 mnuEditUndo.Enabled = False
 txtEdit(1).Visible = False
 txtEdit(0).Visible = True
 txtEdit(0).Font = dlgEdit.FontName
 txtEdit(0).FontSize = dlgEdit.FontSize
 txtEdit(0).Height = Me.ScaleHeight - 10
 txtEdit(0).Width = Me.ScaleWidth - 10
 txtEdit(0) = txtEdit(1)
 txtEdit(1) = ""
End If
End Sub

Private Sub mnuFileExit_Click()
    Unload frmEdit
End Sub

Private Sub mnuFileNew_Click()
    Const conBtns As Integer = vbYesNoCancel + vbExclamation _
                            + vbDefaultButton3 + vbApplicationModal
   Dim conMsg As String
    Dim intUserResponse As Integer
        Me.GetText
    conMsg = "The text in the " & strTitle & " file has changed." & vbCrLf & vbCrLf & "Do you want to save the changes ?"

    If blnChange = True Then        'text box was changed since last save
        intUserResponse = MsgBox(conMsg, conBtns, "Notepadx")
        Select Case intUserResponse
            Case vbYes              'user wants to save current file
                Call mnuFileSave_Click
                If blnCancelSave = True Then
                    Exit Sub
                End If
            Case vbNo               'user does not want to save current file
                                    'process instructions below end if
            Case vbCancel           'user wants to cancel New command
                Exit Sub
        End Select
    End If
    txtEdit(MOST_IMP).Text = ""     'clear text box
    blnChange = False               'reset variable
    frmEdit.Caption = "Untitled - Notepadx"
    dlgEdit.FileName = ""
    
End Sub

Private Sub mnuFileOpen_Click()
 Const conBtns As Integer = vbYesNoCancel + vbExclamation + vbDefaultButton3 + vbApplicationModal
 Const conErr As Integer = vbOKOnly + vbInformation + vbDefaultButton1 + vbApplicationModal
 Const conErrMsg = "This file is too large for Notepadx to open." & vbCrLf & "Please use Windows Wordpad to read this file?"
 Dim conMsg As String
 Dim intUserResponse As Integer
 Dim intUserErr As Integer
 Dim filedata As String
 Dim seeIt As String
 On Error GoTo OpenErrHandler
   dlgEdit.CancelError = True
   Me.GetText
   conMsg = "The text in the " & strTitle & " file has changed." & vbCrLf & vbCrLf & "Do you want to save the changes ?"

 If blnChange = True Then                'document was changed since last save
     intUserResponse = MsgBox(conMsg, conBtns, "Notepadx")
       Select Case intUserResponse
           Case vbYes                      'user wants to save current document
             Call mnuFileSave_Click
              If blnCancelSave = True Then    'user canceled save
                Exit Sub
              End If
            Case vbNo                       'user doesn't want to save current document
                'process instructions below End If
            Case vbCancel                   'user wants to cancel Open command
                Exit Sub
        End Select
    End If
    dlgEdit.Filter = "Text Files(*.txt)|*.txt|All Files(*.*)|*.*"
    dlgEdit.FileName = ""
    dlgEdit.ShowOpen
    
 filedata = FileLen(dlgEdit.FileName)        ' Get information about file size
 
 If Int(filedata) > 65536 Then
  txtEdit(MOST_IMP).Text = ""
  intUserErr = MsgBox(conErrMsg, conErr, "Error")
 Else
  Open dlgEdit.FileName For Input As #1
   txtEdit(MOST_IMP).Text = Input(LOF(1), 1)
    Close #1
    blnChange = False
    If Left(txtEdit(MOST_IMP).Text, 4) = ".LOG" Then
     txtEdit(MOST_IMP).Text = txtEdit(MOST_IMP).Text & vbCrLf & VBA.time & " " & Date
    End If
    frmEdit.Caption = dlgEdit.FileName & " - Notepadx"
    Exit Sub
 End If
OpenErrHandler:
End Sub

Private Sub mnuFilePrint_Click()
    On Error GoTo PrintErrHandler
    dlgEdit.flags = cdlPDNoSelection + cdlPDHidePrintToFile + cdlPDNoPageNums
    dlgEdit.CancelError = True
    dlgEdit.ShowPrinter
    PrintForm
    Exit Sub
    
PrintErrHandler:

End Sub


Private Sub mnuFileSave_Click()
    If frmEdit.Caption = "Untitled - Notepadx" Then
        Call mnuFileSaveAs_Click    'new document
    Else                            'existing document
        Open dlgEdit.FileName For Output As #1
        Print #1, txtEdit(MOST_IMP).Text
        Close #1
        blnChange = False
    End If
End Sub

Private Sub mnuFileSaveAs_Click()
    On Error GoTo SaveErrHandler    'Turn error trapping on
    dlgEdit.CancelError = True      'treat the cancel button as an error
    dlgEdit.flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist
    dlgEdit.Filter = "Text Files(*.txt)|*.txt"
    dlgEdit.ShowSave
    Open dlgEdit.FileName For Output As #1
    Print #1, txtEdit(MOST_IMP).Text
    Close #1
    frmEdit.Caption = dlgEdit.FileName & " - Text Editor"
    blnChange = False
    blnCancelSave = False
    Exit Sub
    
SaveErrHandler:
    blnCancelSave = True            'save was Canceled
    
End Sub



Private Sub mnuEdit_Click()
    If txtEdit(MOST_IMP).SelText = "" Then        'if no text is selected
        mnuEditCut.Enabled = False
        mnuEditCopy.Enabled = False
        mnuEditDelete.Enabled = False
    Else                                'if some text is selected
        mnuEditCut.Enabled = True
        mnuEditCopy.Enabled = True
        mnuEditDelete.Enabled = True
    End If
    If Clipboard.GetText() = "" Then    'if no text is on the clipboard
        mnuEditPaste.Enabled = False
    Else                                'if some text is on the clipboard
        mnuEditPaste.Enabled = True
    End If
End Sub

Private Sub mnuEditCopy_Click()
    Clipboard.Clear                     'clear clipboard
    Clipboard.SetText txtEdit(MOST_IMP).SelText   'send text to clipboard
End Sub


Private Sub mnuEditCut_Click()
    Clipboard.Clear                     'clear clipboard
    Clipboard.SetText txtEdit(MOST_IMP).SelText   'send text to clipboard
    txtEdit(MOST_IMP).SelText = ""                'remove selected text from text box
End Sub

Private Sub mnuEditPaste_Click()
    'retrieve text from clipboard and paste into text box
    txtEdit(MOST_IMP).SelText = Clipboard.GetText()
End Sub
Function GetText()
Dim intFoundPos As Integer
Dim thiss As String
   intFoundPos = InStr(1, frmEdit.Caption, "-", 1)
   thiss = Left(frmEdit.Caption, intFoundPos - 2)
strTitle = thiss
End Function
Private Sub txtEdit_Change(Index As Integer)
If ANOTHER_ONE = 1 Then
 ANOTHER_ONE = 2
 Exit Sub
Else
 blnChange = True
 mnuEditUndo.Enabled = True
 End If
End Sub

Function strSearch(cool As String)
Dim first As String
Dim second As String
Dim third As String
strImp = cool
If strImp = "" Then
 Call mnu23_Click
Else
  With txtEdit(MOST_IMP)
    place = InStr(pos, .Text, cool, vbTextCompare)
    If place > 0 Then
       .SetFocus: .SelStart = place - 1
       .SelLength = Len(cool)
       pos = place + Len(cool)
    Else
       pos = 1
       place = 0
       MsgBox "Cannot find " & "' " & strImp & " '", vbInformation, "Notepadx"
        
    End If
  End With
End If
End Function

Function Another(THIS As String)
    If InStr(txtEdit(MOST_IMP).Text, THIS) <> 0 Then
      txtEdit(MOST_IMP).SetFocus
      txtEdit(MOST_IMP).SelStart = InStr(txtEdit(MOST_IMP), THIS) - 1
      txtEdit(MOST_IMP).SelLength = Len(THIS)
    Else
      MsgBox "Not found"
    End If
End Function

Private Sub txtEdit_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
 Source.Visible = True
End Sub

Private Sub txtEdit_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
 Const conBtns As Integer = vbYesNoCancel + vbExclamation + vbDefaultButton3 + vbApplicationModal
 Const conErr As Integer = vbOKOnly + vbInformation + vbDefaultButton1 + vbApplicationModal
 Const conErrMsg = "This file is too large for Notepadx to open." & vbCrLf & "Please use Windows Wordpad to read this file?"
 Dim conMsg As String
 Dim intUserResponse As Integer
 Dim intUserErr As Integer
 Dim filedata As String
 Dim seeIt As String
 On Error GoTo OpenErrHandler
 
   Me.GetText
   conMsg = "The text in the " & strTitle & " file has changed." & vbCrLf & vbCrLf & "Do you want to save the changes ?"

 If blnChange = True Then                'document was changed since last save
     intUserResponse = MsgBox(conMsg, conBtns, "Notepadx")
       Select Case intUserResponse
           Case vbYes                      'user wants to save current document
             Call mnuFileSave_Click
              If blnCancelSave = True Then    'user canceled save
                Exit Sub
              End If
            Case vbNo                       'user doesn't want to save current document
                'process instructions below End If
            Case vbCancel                   'user wants to cancel Open command
                Exit Sub
        End Select
    End If
    
 filedata = FileLen(Data.Files.Item(1))        ' Get information about file size
 
 If Int(filedata) > 65536 Then
  txtEdit(MOST_IMP).Text = ""
  intUserErr = MsgBox(conErrMsg, conErr, "Error")
 Else
  Open Data.Files.Item(1) For Input As #1
   txtEdit(MOST_IMP).Text = Input(LOF(1), 1)
    Close #1
    blnChange = False
    If Left(txtEdit(MOST_IMP).Text, 4) = ".LOG" Then
     txtEdit(MOST_IMP).Text = txtEdit(MOST_IMP).Text & vbCrLf & VBA.time & " " & Date
    End If
    frmEdit.Caption = Data.Files.Item(1) & " - Notepadx"
    Exit Sub
 End If
OpenErrHandler:
End Sub
