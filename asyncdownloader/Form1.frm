VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar prgBAR 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdDOWNLOAD 
      Caption         =   "Download"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
   Begin Project1.Downloader dwnLOADER 
      Left            =   2280
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin MSComctlLib.ProgressBar prgBAR 
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblLABEL 
      Caption         =   " "
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   4335
   End
   Begin VB.Label lblLABEL 
      Caption         =   " "
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
                (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
                ByVal lParam As Long) As Long
Private Const WM_USER = &H400
Private Const PBM_SETBARCOLOR = (WM_USER + 9)
'------------------------------------------------------------
' Changes to color of a progress bar
'------------------------------------------------------------
Private Sub ChangeProgressBarColor(objProgressBar As ProgressBar, mlngColor As Long)
    Call SendMessageLong(objProgressBar.hwnd, PBM_SETBARCOLOR, 0&, mlngColor)
End Sub
Private Sub cmdDOWNLOAD_Click()
    '------------------------------------------------------------
    ' Show how that the single control and handle multiple
    ' downloads.
    '------------------------------------------------------------
    'download IE SP1
    Me.dwnLOADER.BeginDownload "http://download.microsoft.com/download/ie6sp1/finrel/6_sp1/W98NT42KMeXP/EN-US/ie6setup.exe", App.Path & "\IE6_SETUP.EXE"
    'download DirectX 9.0 Runtime
    Me.dwnLOADER.BeginDownload "http://download.microsoft.com/download/2/2/3/22371837-c4dc-4f8b-af21-00c80d8b235c/dxwebsetup.exe", App.Path & "\dxwebsetup.exe"
End Sub
Private Sub dwnLOADER_DownloadComplete(MaxBytes As Long, SaveFile As String)
    '------------------------------------------------------------
    ' Just showing how to use the SaveFile argument
    ' in determining which file completed.
    '------------------------------------------------------------
    If InStr(SaveFile, "dxwebsetup") Then
        Me.lblLABEL(1).Caption = "Done"
        ChangeProgressBarColor Me.prgBAR(1), vbGreen
    Else
        Me.lblLABEL(0).Caption = "Done"
        ChangeProgressBarColor Me.prgBAR(0), vbGreen
    End If
End Sub
Private Sub dwnLOADER_DownloadProgress(CurBytes As Long, MaxBytes As Long, SaveFile As String)
    '------------------------------------------------------------
    ' Shows how to give feedback to the user and how
    ' to use the SaveFile argument to determine which
    ' file we are talking about.
    '------------------------------------------------------------
    If InStr(SaveFile, "dxwebsetup") Then
        With Me.prgBAR(1)
            .Max = MaxBytes
            .Value = CurBytes
            Me.lblLABEL(1).Caption = CurBytes & " of " & MaxBytes & "..."
        End With
    Else
        With Me.prgBAR(0)
            .Max = MaxBytes
            .Value = CurBytes
            Me.lblLABEL(0).Caption = CurBytes & " of " & MaxBytes & "..."
        End With
    End If
End Sub

Private Sub Form_Load()
    ChangeProgressBarColor Me.prgBAR(0), vbMagenta
    ChangeProgressBarColor Me.prgBAR(1), vbMagenta
End Sub
