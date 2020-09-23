VERSION 5.00
Begin VB.Form frmQuickTip 
   Caption         =   "QuickTip"
   ClientHeight    =   3315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3825
   Icon            =   "frmQuickTip.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3315
   ScaleWidth      =   3825
   StartUpPosition =   3  'Windows Default
   Begin ObjectDoc.QuickTip QuickTip1 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   5741
   End
End
Attribute VB_Name = "frmQuickTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()

    On Error Resume Next
        QuickTip1.Width = Me.ScaleWidth
        QuickTip1.Height = Me.ScaleHeight

End Sub
