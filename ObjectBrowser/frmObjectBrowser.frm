VERSION 5.00
Begin VB.Form frmObjectBrowser 
   Caption         =   "Object browser / ActiveHelp"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4230
   Icon            =   "frmObjectBrowser.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3630
   ScaleWidth      =   4230
   StartUpPosition =   3  'Windows Default
   Begin ObjectDoc.ObjectBrowser ObjectBrowser1 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5741
   End
End
Attribute VB_Name = "frmObjectBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()

    On Error Resume Next
        ObjectBrowser1.Height = Me.ScaleHeight
        ObjectBrowser1.Width = Me.ScaleWidth

End Sub

'Purpose:
' Show the active help for the selected object in the ObjectBrowser

Private Sub ObjectBrowser1_NodeSelect(strNode As String, strCall As String, strDescription As String, strHelp As String, strHelpFile As String, Node As MSComctlLib.Node, strHelpPage As String)

Dim tmpPage() As String
Dim strType As String
Dim strPage As String

    frmQuickTip.QuickTip1.WaitCount = 2
    frmQuickTip.QuickTip1.HelpType = "ObjectBrowser"
    frmQuickTip.QuickTip1.QuickTipKey = strCall

    ' The descriptin and helptekst from the type library
    If Len(strDescription) > 0 Then
        If Len(strHelp) > 0 Then
            frmQuickTip.QuickTip1.HelpText = strDescription & "<BR><BR>" & strHelp
        Else
            frmQuickTip.QuickTip1.HelpText = strDescription
        End If
    Else
        frmQuickTip.QuickTip1.HelpText = strHelp
    End If

    ' If this help page exist then we will use it.
    frmQuickTip.QuickTip1.HelpFile = strHelpPage

End Sub

Private Sub ObjectBrowser1_NodeExecute(strNode As String, strCall As String, strDescription As String, strHelp As String, strHelpFile As String, Node As MSComctlLib.Node, strHelpPage As String)

    MsgBox "You have selected " & vbCrLf & strCall, vbInformation

End Sub

Private Sub ObjectBrowser1_Progress(strCurrentNode As String)

    frmProgress.lblProgress = strCurrentNode
    frmProgress.Refresh

End Sub
