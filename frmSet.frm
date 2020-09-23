VERSION 5.00
Begin VB.Form frmSet 
   Caption         =   "Form1"
   ClientHeight    =   1395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3435
   LinkTopic       =   "Form1"
   ScaleHeight     =   1395
   ScaleWidth      =   3435
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option3 
      Caption         =   "Hours"
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   720
      Width           =   855
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Minutes"
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   480
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Seconds"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   240
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.TextBox txtSet 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Text            =   "0"
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Set Timed Refresh"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Incrament"
      Height          =   1215
      Left            =   2040
      TabIndex        =   4
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Refresh every"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command2_Click()
    Dim TT As Single
    If Option1.Value = True Then
        TT = Val(txtSet.Text)
    ElseIf Option2.Value = True Then
        TT = Val(txtSet.Text) * 60
    ElseIf Option3.Value = True Then
        TT = Val(txtSet.Text) * 3600
    End If
    frmE.sngIntervel = TT
    frmE.cmdST.Enabled = True
    Unload frmSet
End Sub
