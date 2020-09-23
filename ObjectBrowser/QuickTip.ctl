VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.UserControl QuickTip 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4980
   ScaleHeight     =   3360
   ScaleWidth      =   4980
   ToolboxBitmap   =   "QuickTip.ctx":0000
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3840
      Top             =   120
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      ExtentX         =   6376
      ExtentY         =   5106
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
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
      Location        =   ""
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser2 
      Height          =   1095
      Left            =   3360
      TabIndex        =   1
      Top             =   600
      Width           =   975
      ExtentX         =   1720
      ExtentY         =   1931
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
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
      Location        =   ""
   End
End
Attribute VB_Name = "QuickTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' QuickTip control 1.1
'
' Copyright 2005, E.V.I.C.T. B.V.
' Website: http:\\www.evict.nl
' Support: mailto:evict@vermeer.nl
'
'Purpose:
' This control adds a verry nice help function to your application. There is even
' support for getting the help text out of a chm help file. Just look how it works
' when you add one of my DLLs to the object browser (click on various nodes for
' getting the help). Also have a look at ComDoc. That is an application that I build
' for generating chm help files.
'
'License:
' GPL - The GNU General Public License
' Permits anyone the right to use and modify the software without limitations
' as long as proper credits are given and the original and modified source code
' are included. Requires that the final product, software derivate from the
' original source or any software utilizing a GPL component, such as this,
' is also licensed under the GPL license.
' For more information see http://www.gnu.org/licenses/gpl.txt
'
'License adition:
' You are permitted to use the software in a non-commercial context free of
' charge as long as proper credits are given and the original unmodified source
' code is included.
' For more information see http://www.evict.nl/licenses.html
'
'License exeption:
' If you would like to obtain a commercial license then please contact E.V.I.C.T. B.V.
' For more information see http://www.evict.nl/licenses.html
'
'Terms:
' This software is provided "as is", without warranty of any kind, express or
' implied, including  but not limited to the warranties of merchantability,
' fitness for a particular purpose and noninfringement. In no event shall the
' authors or copyright holders be liable for any claim, damages or other
' liability, whether in an action of contract, tort or otherwise, arising
' from, out of or in connection with the software or the use or other
' dealings in the software.
'
'History:
' 2002 : Created and added to the sharware library siteskinner
' jan 2005 : Changed the licensing from shareware to opensource
' feb 2005 : Added the SetSockLinger and getascip for the Improved connection method

Option Explicit

Dim piWaitCount As Integer
Dim psQuickTipKey As String
Dim psHelpType As String
Dim psHelpFile As String
Dim psHelpText As String
Dim strAH() As String

Private Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Long) As Long
Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Purpose:
' With this you can delay the display of a help with WaitCount/10 seconds. You have to set this before setting the QuickTipKey.

Public Property Get WaitCount() As Integer

    WaitCount = piWaitCount

End Property

Public Property Let WaitCount(NewWaitCount As Integer)

    piWaitCount = NewWaitCount

End Property

'Purpose:
' This is the key in the Active help file

Public Property Get QuickTipKey() As String

    QuickTipKey = psQuickTipKey

End Property

Public Property Let QuickTipKey(NewQuickTipKey As String)

    psQuickTipKey = NewQuickTipKey

End Property

'Purpose:
' This is the type of help

Public Property Get HelpType() As String

    HelpType = psHelpType

End Property

Public Property Let HelpType(NewHelpType As String)

    psHelpType = NewHelpType

End Property

'Purpose:
' This is the Help file

Public Property Get HelpFile() As String

    HelpFile = psHelpFile

End Property

Public Property Let HelpFile(NewHelpFile As String)

    psHelpFile = NewHelpFile

End Property

'Purpose:
' This is the Help text

Public Property Get HelpText() As String

    HelpText = psHelpText

End Property

Public Property Let HelpText(NewHelpText As String)

    psHelpText = NewHelpText

End Property

'Purpose:
' this will parse a file into a 2 dimentional array using the pipe simbol as column seperator and a return as the row seperator.

Public Sub LoadQuickTip(strFileName As String)

    On Error GoTo exitthis
Dim strAHt() As String
Dim strAHt2() As String
Dim I As Integer
Dim j As Integer

    strAHt2 = Split(LoadFile(strFileName), vbCrLf)
    ReDim strAH(UBound(strAHt2), 4) As String
    For I = 0 To UBound(strAH) - 1
        strAHt = Split(strAHt2(I), "|")
        For j = 0 To UBound(strAHt) - 1
            strAH(I, j) = strAHt(j)
        Next j
    Next I
exitthis:

End Sub

Private Sub UserControl_Initialize()

    WebBrowser1.Navigate "about:blank"

End Sub

Private Sub UserControl_Resize()

    WebBrowser1.Width = UserControl.Width
    WebBrowser1.Height = UserControl.Height

End Sub

Private Sub WebBrowser1_BeforeNavigate2(ByVal pDisp As Object, url As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)

' In these cases we just navigate

    If url = "about:blank" Then Exit Sub
    If Left$(url, 4) = "mk:@" Then Exit Sub

    ' Cancel the navigation and build the active help
    piWaitCount = 0
    psQuickTipKey = Mid$(url, 10)
    psQuickTipKey = Replace(psQuickTipKey, "%20", " ")
    psHelpType = Trim$(Mid$(psQuickTipKey & "   ", InStr(1, psQuickTipKey & "?", "?") + 1))
    psQuickTipKey = Left$(psQuickTipKey, InStr(1, psQuickTipKey & "?", "?") - 1)
    ShowQuickTip
    Cancel = True

End Sub

Private Sub Timer1_Timer()

' Used for delaid help

    If piWaitCount = 1 Then
        piWaitCount = 0
        ShowQuickTip
    Else
        If piWaitCount > 1 Then
            piWaitCount = piWaitCount - 1
        End If
    End If

End Sub

Private Sub ShowQuickTip()

Dim strHTML As String
Dim strTemp As String

'General settings

    strHTML = "<span style='FONT-SIZE: 8pt; MARGIN: 0px; COLOR: #1687f8; FONT-FAMILY: Verdana, Arial, Helvetica'>"
    ' Title
    strHTML = strHTML & "<b>" & psQuickTipKey & "</b><BR><BR>"

    Select Case psHelpType
    Case "Interface"
        strHTML = strHTML & GetHelp(psQuickTipKey)
    Case "HtmlHelp"
        ' This is used for a navigation inside the webbrowser control to navigate to a page inside the chm file
        RunThisURL psHelpFile
        Exit Sub
    Case "ObjectBrowser"
        strTemp = GetDescription(psHelpFile)
        If strTemp = "" Then
            strHTML = strHTML & psHelpText
        Else
            strHTML = strHTML & strTemp & "<BR><BR>" & "For more informaition click <a href=""" & psHelpFile & """ target=_new>here</a>"
        End If
    Case Else
    End Select
    ' Now display the help.
    WebBrowser1.Document.Body.innerHTML = strHTML & "</span>"

End Sub

'Purpose
' Get the description from the chm file

Private Function GetDescription(strHelpFile As String) As String

    On Error GoTo exitfunc
Dim X As String
Dim I As Integer

    WebBrowser2.Navigate strHelpFile
    While WebBrowser2.ReadyState <> READYSTATE_COMPLETE And WebBrowser2.ReadyState <> READYSTATE_INTERACTIVE
        DoEvents
    Wend
    X = WebBrowser2.Document.Body.outerHTML
    I = InStr(1, X, "<!-- QuickTipStart -->")
    If I > 0 Then
        GetDescription = Mid$(X, I, InStr(1, X, "<!-- QuickTipEnd -->") - I)
    Else
        GetDescription = ""
    End If

exitfunc:

End Function

'Purpose:
' Get the helpteskst for a given key

Private Function GetHelp(strHelpKey As String) As String

    On Error GoTo exitthis
Dim I As Integer

    For I = 0 To UBound(strAH) - 1
        If strAH(I, 0) = strHelpKey Then
            GetHelp = strAH(I, 1)
            Exit Function
        End If
    Next I
exitthis:

End Function

'Purpose:
' This function will load a file and return the data in a string.

Private Function LoadFile(FileName As String) As String

Dim intnextfreefile As Integer

    intnextfreefile = FreeFile
    Open FileName For Binary As #intnextfreefile
    LoadFile = String(LOF(intnextfreefile), vbNullChar) 'must initialize the variable
    Get #intnextfreefile, , LoadFile
    Close #intnextfreefile

End Function

'Purpose:
' This functions will open any URL with your default web browser.

Private Sub RunThisURL(strURL As String)

Dim strFileName    As String
Dim strDummy       As String
Dim strBrowserExec As String * 255
Dim lngRetVal      As Long
Dim intFileNumber  As Integer

' Create a temporary HTM file

    strBrowserExec = Space(255)
    strFileName = "~TempBrowserCheck.HTM"
    intFileNumber = FreeFile
    Open strFileName For Output As #intFileNumber
    Write #intFileNumber, "<HTML> <\HTML>"
    Close #intFileNumber

    ' Find the default browser.
    lngRetVal = FindExecutable(strFileName, strDummy, strBrowserExec)
    strBrowserExec = Trim$(strBrowserExec)

    ' If an application is found, launch it!
    If lngRetVal <= 32 Or IsEmpty(strBrowserExec) Then
        MsgBox "Could not find your Browser", vbExclamation, "Browser Not Found"
    Else
        lngRetVal = ShellExecute(App.hInstance, "open", strBrowserExec, strURL, strDummy, 1)
        If lngRetVal <= 32 Then
            MsgBox "Web Page not Opened", vbExclamation, "URL Failed"
        End If
    End If

    ' remove the temporary file
    Kill strFileName

Exit Sub

End Sub
