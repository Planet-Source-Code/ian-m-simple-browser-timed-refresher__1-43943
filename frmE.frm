VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmE 
   Caption         =   "eBay Browser"
   ClientHeight    =   9675
   ClientLeft      =   5130
   ClientTop       =   2415
   ClientWidth     =   12600
   LinkTopic       =   "Form1"
   ScaleHeight     =   9675
   ScaleWidth      =   12600
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "refresh"
      Height          =   495
      Left            =   6960
      TabIndex        =   8
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdFor 
      Caption         =   "Forward"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "eBay"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdST 
      Caption         =   "Start Refresher"
      Enabled         =   0   'False
      Height          =   495
      Left            =   10920
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   12000
      Top             =   120
   End
   Begin SHDocVwCtl.WebBrowser Web 
      Height          =   8295
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   12255
      ExtentX         =   21616
      ExtentY         =   14631
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&GO"
      Default         =   -1  'True
      Height          =   255
      Left            =   10080
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtURL 
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Text            =   "http://www.ebay.com/"
      Top             =   120
      Width           =   9495
   End
   Begin VB.CommandButton cmdET 
      Caption         =   "Stop  Refresher"
      Height          =   495
      Left            =   10920
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu edit 
      Caption         =   "&Edit"
      Begin VB.Menu re 
         Caption         =   "Refresh setings..."
      End
   End
End
Attribute VB_Name = "frmE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'progamer: Ian McCall
'
Public sngIntervel As Single
Dim intBack As Integer, Time As Single
Option Explicit

Private Sub cmdBack_Click()
    Web.GoBack
    cmdFor.Enabled = True
    intBack = intBack + 1
End Sub

Private Sub Command1_Click()
    Web.Navigate (txtURL.Text)
    cmdFor.Enabled = False
    intBack = 0
End Sub

Private Sub Command2_Click()
    Web.Refresh
End Sub

Private Sub Command3_Click()
    Web.Navigate ("http://www.ebay.com/")
End Sub

Private Sub cmdET_Click()
    Timer1.Enabled = False
    cmdET.Visible = False
    cmdST.Visible = True
    Time = 0
End Sub

Private Sub cmdST_Click()
    Timer1.Enabled = True
    cmdET.Visible = True
    cmdST.Visible = False
End Sub

Private Sub cmdFor_Click()
    Web.GoForward
    intBack = intBack - 1
    If intBack = 0 Then
        cmdFor.Enabled = False
    End If
End Sub

Private Sub exit_Click()
    End
End Sub

Private Sub Form_Load()
    Web.Navigate (txtURL.Text)
End Sub

Private Sub Form_Resize()
    If frmE.WindowState <> 1 Then
        Web.Width = frmE.ScaleWidth
        Web.Height = frmE.ScaleHeight - 750
    End If
End Sub

Private Sub re_Click()
    Load frmSet
    frmSet.Visible = True
End Sub

Private Sub Timer1_Timer()
    Time = Time + 1
    If Time = sngIntervel Then
        Web.Refresh
        Time = 0
    End If
End Sub

Private Sub Web_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    txtURL.Text = Web.LocationURL
    If Web.LocationURL <> "http://www.ebay.com/" Then
        cmdBack.Enabled = True
    Else
        cmdBack.Enabled = False
    End If
End Sub

