VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4695
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   ScaleHeight     =   313
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   504
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   " Rayment's Bit Operations "
      Height          =   1815
      Left            =   4080
      TabIndex        =   8
      Top             =   2640
      Width           =   3375
      Begin VB.CommandButton Command2 
         Caption         =   "Test OS Modified"
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   1320
         Width           =   2655
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Test XP Only"
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label ModLbl 
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Label RayBitLbl 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " CptnVic's All OS Bit Operations "
      Height          =   1095
      Left            =   4080
      TabIndex        =   5
      Top             =   1440
      Width           =   3375
      Begin VB.CommandButton Command3 
         Caption         =   "Test Cptn Vic's Bits"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label VicBitLbl 
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3135
      End
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   6960
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.bmp;*.gif;*.jpg"
   End
   Begin VB.Frame Frame1 
      Caption         =   " Long Stream GetDIBits Module Test "
      Height          =   1215
      Left            =   4080
      TabIndex        =   1
      Top             =   120
      Width           =   3375
      Begin VB.CommandButton Command1 
         Caption         =   "Test Long Stream"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label LSTimeLbl 
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   120
      ScaleHeight     =   289
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   257
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.PictureBox testPic 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   11520
         Left            =   0
         ScaleHeight     =   768
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1024
         TabIndex        =   3
         Top             =   240
         Width           =   15360
      End
   End
   Begin VB.Menu mnuFileMain 
      Caption         =   "File"
      Begin VB.Menu mnuLoadPic 
         Caption         =   "Load Picture"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TickIn As Double, TickOut As Double, TickTtl As Double 'For time calculations only
Private Sub Command1_Click()
    Dim Q As Long
    Me.MousePointer = 11 'show a wait coming
    TickIn = GetTickCount& 'get the starting tick count
    Q = UniqColors(testPic) 'count the colors
    TickOut = GetTickCount& 'get ending tick count
    TickTtl = TickOut - TickIn 'calc time used
    Me.MousePointer = 0 ' reset the mouse pointer
    MsgBox "Unique Colors = " & Q & vbCrLf & vbCrLf & "That took " & Format(TickTtl / 1000, "####.####") & " Secs.", vbOKOnly, "Test #1 Finished"
    LSTimeLbl.Caption = "Colors: " & Q & "  Time: " & Format(TickTtl / 1000, "####.####") & " Secs."
End Sub

Private Sub Command2_Click()
    Dim Q As Long
    Me.MousePointer = 11 'show a wait coming
    TickIn = GetTickCount& 'get the starting tick count
    Q = ColorCountBitsNew(testPic) 'count the colors
    TickOut = GetTickCount& 'get ending tick count
    TickTtl = TickOut - TickIn 'calc time used
    Me.MousePointer = 0 ' reset the mouse pointer
    MsgBox "Unique Colors = " & Q & vbCrLf & vbCrLf & "That took " & Format(TickTtl / 1000, "####.####") & " Secs.", vbOKOnly, "Test #1 Finished"
    ModLbl.Caption = "Colors: " & Q & "  Time: " & Format(TickTtl / 1000, "####.####") & " Secs."

End Sub

Private Sub Command3_Click()
    Dim Q As Long
    Me.MousePointer = 11 'show a wait coming
    TickIn = GetTickCount& 'get the starting tick count
    Q = UniqBitColors(testPic) 'count the colors
    TickOut = GetTickCount& 'get ending tick count
    TickTtl = TickOut - TickIn 'calc time used
    Me.MousePointer = 0 ' reset the mouse pointer
    MsgBox "Unique Colors = " & Q & vbCrLf & vbCrLf & "That took " & Format(TickTtl / 1000, "####.####") & " Secs.", vbOKOnly, "Test #1 Finished"
    VicBitLbl.Caption = "Colors: " & Q & "  Time: " & Format(TickTtl / 1000, "####.####") & " Secs."
End Sub

Private Sub Command4_Click()
    Dim Q As Long
    Me.MousePointer = 11 'show a wait coming
    TickIn = GetTickCount& 'get the starting tick count
    Q = ColorCountBits(testPic) 'count the colors
    TickOut = GetTickCount& 'get ending tick count
    TickTtl = TickOut - TickIn 'calc time used
    Me.MousePointer = 0 ' reset the mouse pointer
    MsgBox "Unique Colors = " & Q & vbCrLf & vbCrLf & "That took " & Format(TickTtl / 1000, "####.####") & " Secs.", vbOKOnly, "Test #1 Finished"
    RayBitLbl.Caption = "Colors: " & Q & "  Time: " & Format(TickTtl / 1000, "####.####") & " Secs."
End Sub

Private Sub Form_Load()
    With testPic
        .Left = 0
        .Top = 0
    End With
    Me.Caption = "Use the menu to load a picture."
End Sub

Private Sub mnuLoadPic_Click()
    CDialog.FileName = ""
    CDialog.ShowOpen
    If CDialog.FileName <> "" Then
        testPic.Picture = LoadPicture(CDialog.FileName)
        testPic.Refresh
        Me.Caption = testPic.Width & " px X " & testPic.Height & " px = " & testPic.Width * testPic.Height & " pixels to check"
        LSTimeLbl.Caption = ""
        VicBitLbl.Caption = ""
        RayBitLbl.Caption = ""
        ModLbl.Caption = ""
    End If
End Sub
