VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmControlPanel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Component Documenter Control Pannel"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCustomObject 
      Caption         =   "Add a custom object to the browser"
      Height          =   495
      Left            =   2880
      TabIndex        =   11
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Caption         =   "Settings used for add to browser"
      Height          =   2415
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   2655
      Begin VB.CheckBox settings 
         Caption         =   "Group by type"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1695
      End
      Begin VB.CheckBox settings 
         Caption         =   "Enum objects in root"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   1935
      End
      Begin VB.CheckBox settings 
         Caption         =   "Extend inner objects"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CheckBox settings 
         Caption         =   "Extend parent properties"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CheckBox settings 
         Caption         =   "Remove  refered objects"
         Height          =   255
         Index           =   16
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdFileOpen 
      Caption         =   "..."
      Height          =   255
      Left            =   5160
      TabIndex        =   2
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cmdFileBrowse 
      Caption         =   "add to browser"
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   1920
      Width           =   2655
   End
   Begin VB.TextBox txtObject 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   5055
   End
   Begin MSComDlg.CommonDialog cdlTypeLibrary 
      Left            =   5160
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Object browser"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label lblTip 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hoover your mouse over any control to see extra information in the QuickTip window"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   615
      Left            =   0
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   10
      Top             =   0
      Width           =   5655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "We will process this file:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   3975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000018&
      BackStyle       =   1  'Opaque
      Height          =   3135
      Left            =   0
      Top             =   1320
      Width           =   5655
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   735
      Left            =   0
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   13
      Top             =   600
      Width           =   5655
   End
End
Attribute VB_Name = "frmControlPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strTemplate As String

Dim strPublication As String


Private Sub DBPublisher_Status(strStatus As String)

    Debug.Print strStatus

End Sub

'Purpose:
' Initialize some things

Private Sub Form_Load()

Dim strTemplate As String

    Me.Show

    'Make sure we have all the help teksts
    frmQuickTip.QuickTip1.LoadQuickTip (App.Path & "\QuickTip.dat")

    ' This is what can be used in the <a href="QuickTip:tipname"> tags in the QuickTip.dat
    frmQuickTip.QuickTip1.QuickTipKey = "QuickTip"

    ' The default settings
    settings(OB_ExtendInnerObjects) = 1
    settings(OB_RemoveReferedObjects) = 1

    ' Just show the name of the SSPro dll
    If Dir("C:\Program Files\Site Skinner Pro\siteSkinnerDev.dll") <> "" Then
        Me.txtObject = "C:\Program Files\Site Skinner Pro\SiteSkinnerDev.dll"
    Else
        Me.txtObject = "SiteSkinnerDev.dll"
    End If

    frmQuickTip.Show

End Sub

'Purpose:
' Make sure everything is closed

Private Sub Form_Unload(Cancel As Integer)

    Unload frmObjectBrowser
    Unload frmProgress
    Unload frmQuickTip

End Sub

'Purpose:
' Solicit the user for a component that also has a type library

Private Sub cmdFileOpen_Click()

    cdlTypeLibrary.DialogTitle = "Open Type Library"
    cdlTypeLibrary.InitDir = "d:\winnt\system32\"
    cdlTypeLibrary.Filter = "Type Libraries (*.tlb;*.olb;*.dll;*.ocx)|*.tlb;*.olb;*.dll;*.ocx|All Files (*.*)|*.*"
    cdlTypeLibrary.ShowOpen
    If Err.Number <> cdlCancel Then Me.txtObject = cdlTypeLibrary.FileName

End Sub

'Purpose:
' Add the selected component to the object browser

Private Sub cmdFileBrowse_Click()

' Add the Defautlt Visual basic libraries
'ObjectBrowser1.settings = OB_ExtendInnerObjects
'   ObjectBrowser1.AddFromFile ("vb6.olb")
'   ObjectBrowser1.AddFromFile ("msvbvm60.dll\1")
'   ObjectBrowser1.AddFromFile ("msvbvm60.dll\3")
'   ObjectBrowser1.AddFromFile ("stdole2.tlb")
'   ObjectBrowser1.AddFromFile ("tlbinf32.dll")

' Add the Defautlt MS Office
'ObjectBrowser1.settings = 0 ' They are to big for extending inner objects
'   ObjectBrowser1.AddFromFile ("Excel9.olb") ' For Excel 200
'   ObjectBrowser1.AddFromFile ("Excel.olb")  ' For Excel 2002
'   ObjectBrowser1.AddFromFile ("Mso.olb")    ' For Office XP
'   See also MSWord9.olb, MSOutl9.olb, MSAcc9.olb, Graph9.olb, MSBdr9.olb, MSWord.olb, MSOUTL.olb, MSPPT.olb, MSAcc.olb, Graph.olb
'ObjectBrowser1.settings = OB_ExtendInnerObjects

    frmObjectBrowser.Show
    Screen.MousePointer = vbHourglass
    frmProgress.Show
    frmObjectBrowser.ObjectBrowser1.settings = OB_GroupMemberType * settings(OB_GroupMemberType) + _
                                               OB_EnumInRoot * settings(OB_EnumInRoot) + _
                                               OB_ExtendInnerObjects * settings(OB_ExtendInnerObjects) + _
                                               OB_ExtendParent * settings(OB_ExtendParent) + _
                                               OB_RemoveReferedObjects * settings(OB_RemoveReferedObjects)
    If frmObjectBrowser.ObjectBrowser1.AddFromFile(Me.txtObject) <> 0 Then
        MsgBox ("You did not select a valid file. The file should be a OLE/COM component that has a type library.")
    End If
    frmProgress.Hide
    Screen.MousePointer = vbNormal

End Sub

'Purpose:
' I used this in SiteSkinner for adding a custom tree for the internal objects that were made available to the scripting control.

Private Sub cmdCustomObject_Click()

Dim nodx As Node

    frmObjectBrowser.Show
    With frmObjectBrowser.ObjectBrowser1.Nodes
        Set nodx = .Add(, , "SiteSkinnerPro", "SiteSkinnerPro", 10, 10)
        nodx.Sorted = True

        Set nodx = .Add("SiteSkinnerPro", tvwChild, "SiteSkinnerPro.Application", "Application", 1, 1)
        nodx.Sorted = True

        Set nodx = .Add("SiteSkinnerPro", tvwChild, "SiteSkinnerPro.ScriptControl", "ScriptControl", 11, 11)
        nodx.Sorted = True
        frmObjectBrowser.ObjectBrowser1.AddForeignMembersToNode nodx, "msscript.ocx", "ScriptControl"

        Set nodx = .Add("SiteSkinnerPro", tvwChild, "SiteSkinnerPro.MainForm", "MainForm", 12, 12)
        nodx.Sorted = True
        frmObjectBrowser.ObjectBrowser1.AddForeignMembersToNode nodx, "vb6.olb", "Form"

        Set nodx = .Add("SiteSkinnerPro", tvwChild, "SiteSkinnerPro.App", "App", 1, 1)
        nodx.Sorted = True
        frmObjectBrowser.ObjectBrowser1.AddForeignMembersToNode nodx, "vb6.olb", "App"

        Set nodx = .Add("SiteSkinnerPro", tvwChild, "SiteSkinnerPro.Screen", "Screen", 1, 1)
        nodx.Sorted = True
        frmObjectBrowser.ObjectBrowser1.AddForeignMembersToNode nodx, "vb6.olb", "Screen"

        Set nodx = .Add("SiteSkinnerPro", tvwChild, "SiteSkinnerPro.Script", "Script", 13, 13)
        nodx.Sorted = True

        Set nodx = .Add("SiteSkinnerPro.Application", tvwChild, "SiteSkinnerPro.Application.Method", "Methods", 7, 7)
        nodx.Sorted = True

        nodx.EnsureVisible

        .Add "SiteSkinnerPro.Application.Method", tvwChild, "", "Show", 7, 7
        .Add "SiteSkinnerPro.Application.Method", tvwChild, "", "Hide", 7, 7
        .Add "SiteSkinnerPro.Application.Method", tvwChild, "", "LoadFile strFileName", 7, 7
        .Add "SiteSkinnerPro.Application.Method", tvwChild, "", "ToolExecute strToolName", 7, 7

    End With

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    QT 20, "Interface", "Control pannel"

End Sub

Private Sub cmdFileOpen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    QT 20, "Interface", "Open a file"

End Sub

Private Sub settings_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    QT 20, "Interface", "ObjectBrowser settings"

End Sub

Private Sub txtObject_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    QT 20, "Interface", "Selected file"

End Sub

Private Sub cmdCustomObject_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    QT 20, "Interface", "Custom objects"

End Sub

Private Sub lblTip_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    QT 20, "Interface", "QuickTip"

End Sub

Private Sub cmdFileBrowse_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    QT 20, "Interface", "File to browser"

End Sub

'Purpose:
' Quick function for the QuickTip

Private Sub QT(intCount As Integer, strType As String, strKey As String)

    frmQuickTip.QuickTip1.WaitCount = intCount
    frmQuickTip.QuickTip1.HelpType = strType
    frmQuickTip.QuickTip1.QuickTipKey = strKey

End Sub
