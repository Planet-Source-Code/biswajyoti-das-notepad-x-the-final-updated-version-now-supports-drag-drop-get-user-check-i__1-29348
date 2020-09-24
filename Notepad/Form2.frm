VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Notepadx"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   ClipControls    =   0   'False
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   135
      Left            =   1080
      TabIndex        =   6
      Top             =   2400
      Width           =   3855
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   120
      Picture         =   "Form2.frx":000C
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   0
      Top             =   240
      Width           =   750
   End
   Begin VB.Label lblRec 
      Caption         =   "Label6"
      Height          =   375
      Left            =   4200
      TabIndex        =   10
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Available Physical Memory"
      Height          =   255
      Left            =   1080
      TabIndex        =   9
      Top             =   3120
      Width           =   3015
   End
   Begin VB.Label lblMem 
      Caption         =   "Label5"
      Height          =   255
      Left            =   4200
      TabIndex        =   8
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Total Physical Memory: "
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Label lblNameOrg 
      Caption         =   "Label4"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   2040
      Width           =   3615
   End
   Begin VB.Label lblName 
      Caption         =   "Label4"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   1680
      Width           =   3615
   End
   Begin VB.Label Label3 
      Caption         =   "This product is distributed freely to:"
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   1320
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "Copyright (C) Biswajyoti Das"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Notepad X"
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   360
      Width           =   3495
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
':-) All for making a Notepad Clone .....
Private Declare Sub GlobalMemoryStatus Lib "kernel32.dll" (lpBuffer As MEMORYSTATUS)
Private Type MEMORYSTATUS
  dwLength As Long
  dwMemoryLoad As Long
  dwTotalPhys As Long
  dwAvailPhys As Long
  dwTotalPageFile As Long
  dwAvailPageFile As Long
  dwTotalVirtual As Long
  dwAvailVirtual As Long
End Type
Private Sub Form_Load()
Dim ms As MEMORYSTATUS
'Get the current memory status.
GlobalMemoryStatus ms
lblMem.Caption = ms.dwTotalPhys \ 1024 & " KB"
lblRec.Caption = ms.dwAvailPhys \ 1024 & "KB"
lblName.Caption = GetStringValue("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "RegisteredOwner")
lblNameOrg.Caption = GetStringValue("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "RegisteredOrganization")
End Sub

Private Sub Picture1_Click()
Form1.Show 1
End Sub
