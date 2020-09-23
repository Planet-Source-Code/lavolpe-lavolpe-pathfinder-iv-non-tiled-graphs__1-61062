VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Drag A/Z Targets with Mouse"
   ClientHeight    =   9360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   624
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   687
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkNoInitPath 
      Caption         =   "Don't show Initial (Red) Path"
      Height          =   195
      Left            =   5490
      TabIndex        =   13
      Top             =   9135
      Width           =   3705
   End
   Begin VB.CheckBox chkFrameOn 
      Caption         =   "Show Regional Rectangles"
      Height          =   195
      Left            =   1530
      TabIndex        =   12
      Top             =   9135
      Width           =   3705
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Hide Path"
      Height          =   630
      Left            =   4020
      TabIndex        =   11
      Top             =   8475
      Width           =   1215
   End
   Begin VB.ComboBox cboBlocks 
      Height          =   315
      ItemData        =   "Pathv4.frx":0000
      Left            =   90
      List            =   "Pathv4.frx":0040
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   8730
      Width           =   1365
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Copy  A/Z"
      Height          =   345
      Left            =   9000
      TabIndex        =   8
      Top             =   8415
      Width           =   1230
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Paste A/Z"
      Height          =   345
      Left            =   9000
      TabIndex        =   7
      Top             =   8775
      Width           =   1230
   End
   Begin VB.TextBox txtPath 
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   5490
      TabIndex        =   6
      Top             =   8835
      Width           =   3405
   End
   Begin VB.TextBox txtPath 
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   0
      Left            =   5490
      TabIndex        =   5
      Top             =   8475
      Width           =   3405
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Find Path"
      Height          =   630
      Left            =   2775
      TabIndex        =   2
      Top             =   8475
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate Graph"
      Height          =   630
      Left            =   1515
      TabIndex        =   1
      Top             =   8475
      Width           =   1215
   End
   Begin VB.PictureBox picGraph 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8265
      Left            =   90
      ScaleHeight     =   551
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   674
      TabIndex        =   0
      Top             =   120
      Width           =   10110
      Begin VB.CommandButton cmdStart 
         BackColor       =   &H000000FF&
         Caption         =   "Z"
         Height          =   240
         Index           =   1
         Left            =   9870
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   240
      End
      Begin VB.CommandButton cmdStart 
         BackColor       =   &H00FFFFFF&
         Caption         =   "A"
         Height          =   240
         Index           =   0
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   8025
         Width           =   240
      End
      Begin VB.Line ln1st 
         BorderColor     =   &H000000C0&
         Index           =   0
         Visible         =   0   'False
         X1              =   287
         X2              =   514
         Y1              =   491
         Y2              =   491
      End
      Begin VB.Line lnPath 
         Index           =   0
         Visible         =   0   'False
         X1              =   287
         X2              =   527
         Y1              =   499
         Y2              =   499
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Number Obstacles"
      Height          =   240
      Left            =   90
      TabIndex        =   10
      Top             =   8490
      Width           =   1350
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' This is simply a sample form to show how the class can be used

' used to drag the A/Z nodes around on the form
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Const WM_NCLBUTTONDOWN As Long = &HA1
Private Const HTCAPTION As Long = 2

' used to create the obstacles
Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateRectRgnIndirect Lib "gdi32.dll" (ByRef lpRect As RECT) As Long

' stopwatch
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long

' draw the graph
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function GetRegionData Lib "gdi32" (ByVal hRgn As Long, ByVal dwCount As Long, lpRgnData As Any) As Long
Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function FillRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private FinalPath() As Long         ' 2-dim array of key points in a path
Private hRgn As Long                ' the region or graph
Private timerStart As Long          ' time test
Private pathCount As Long           ' number of paths found before optimal path finished

' The class has only 1 event and only 2 properties
Private WithEvents pFinder As clsLVpathFinder
Attribute pFinder.VB_VarHelpID = -1

Private Sub chkFrameOn_Click()
    If hRgn Then RedrawRects (chkFrameOn = 1)
End Sub

Private Sub Command2_Click()
' Find the path

If hRgn = 0 Then
    ' didn't generate a graph, so pass a simple region
    With picGraph
        hRgn = CreateRectRgn(0, 0, .Width, .Height)
    End With
End If

ClearPaths
txtPath(0).Text = ""
txtPath(1).Text = ""
RedrawRects (chkFrameOn = 1)
DoEvents

Dim X1 As Long, X2 As Long
Dim Y1 As Long, Y2 As Long

With cmdStart(0)
    X1 = .Left + .Width / 2
    Y1 = .Top + .Height / 2
End With
With cmdStart(1)
    X2 = .Left + .Width / 2
    Y2 = .Top + .Height / 2
End With

If pFinder Is Nothing Then Set pFinder = New clsLVpathFinder
pathCount = 0                   ' reset
timerStart = GetTickCount       ' start timer
pFinder.FindPath X1, Y1, X2, Y2, hRgn   ' find the path
' the results are returned via the pFinder.PathFound event

End Sub

Private Sub showPath(bFinal As Boolean)
' Draws the path using Line controls

Dim nrAnchors As Long
If bFinal Then
    For nrAnchors = 0 To UBound(FinalPath, 2) - 1
        If nrAnchors > lnPath.UBound Then
            Load lnPath(lnPath.UBound + 1)
        End If
        lnPath(nrAnchors).X1 = FinalPath(0, nrAnchors)
        lnPath(nrAnchors).Y1 = FinalPath(1, nrAnchors)

        lnPath(nrAnchors).X2 = FinalPath(0, nrAnchors + 1)
        lnPath(nrAnchors).Y2 = FinalPath(1, nrAnchors + 1)
        lnPath(nrAnchors).Visible = True
        lnPath(nrAnchors).ZOrder
    Next
    For nrAnchors = lnPath.UBound To UBound(FinalPath, 2) Step -1
        Unload lnPath(nrAnchors)
    Next
Else
    For nrAnchors = 0 To UBound(FinalPath, 2) - 1
        If nrAnchors > ln1st.UBound Then
            Load ln1st(ln1st.UBound + 1)
        End If
        ln1st(nrAnchors).X1 = FinalPath(0, nrAnchors)
        ln1st(nrAnchors).Y1 = FinalPath(1, nrAnchors)
        
        ln1st(nrAnchors).X2 = FinalPath(0, nrAnchors + 1)
        ln1st(nrAnchors).Y2 = FinalPath(1, nrAnchors + 1)
        ln1st(nrAnchors).Visible = True
    Next
    For nrAnchors = ln1st.UBound To UBound(FinalPath, 2) Step -1
        Unload ln1st(nrAnchors)
    Next
End If
End Sub

Private Sub Command3_Click()
' used by me only. Allows me to run multiple versions to compare speed tests
' using different heuristic calculations on identical start/end nodes
Dim sClip As String, vClip() As String
sClip = Clipboard.GetText
If Len(sClip) = 0 Then Exit Sub
If Left$(sClip, 7) = "LVPath:" Then
    vClip = Split(sClip, ":")
    cmdStart(0).Move Val(vClip(1)), Val(vClip(2))
    cmdStart(1).Move Val(vClip(3)), Val(vClip(4))
End If
End Sub

Private Sub Command4_Click()
' used by me only. Allows me to run multiple versions to compare speed tests
' using different heuristic calculations on identical start/end nodes
Dim sClip As String
sClip = "LVPath:" & cmdStart(0).Left & ":" & cmdStart(0).Top
sClip = sClip & ":" & cmdStart(1).Left & ":" & cmdStart(1).Top
Clipboard.Clear
Clipboard.SetText sClip
End Sub

Private Sub Command5_Click()
    ClearPaths
End Sub

Private Sub Form_Load()
picGraph.AutoRedraw = True
cboBlocks.ListIndex = 9
'Randomize Timer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If hRgn Then DeleteObject hRgn
    If Not pFinder Is Nothing Then Set pFinder = Nothing
End Sub

Private Sub Command1_Click()

' local routine to simply randomize a graph/map

Dim tRect As RECT
Dim tWidth As Long, tHeight As Long
Dim tRgn As Long, x As Long

With picGraph
    SetRect tRect, 0, 0, .Width, .Height
End With
If hRgn Then DeleteObject hRgn
hRgn = CreateRectRgnIndirect(tRect)
    

For x = 0 To (cboBlocks.ListIndex + 1) * 10

    With picGraph
        tRect.Right = CLng(Rnd * .Width \ 8 + 1)
        tRect.Bottom = CLng(Rnd * .Height \ 8 + 1)
        tRect.Left = CLng(Rnd * (.Width - .Width \ 8)) + .Left - 1
        tRect.Top = CLng(Rnd * (.Height - .Height \ 8)) + .Top - 1
        tRect.Right = tRect.Left + tRect.Right
        tRect.Bottom = tRect.Top + tRect.Bottom
    End With
    
    tRgn = CreateRectRgnIndirect(tRect)
    CombineRgn hRgn, hRgn, tRgn, 4
    DeleteObject tRgn

Next

Call Command2_Click
End Sub

Private Sub cmdStart_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
' Using A,E,F,C keys you can move the A/Z nodes around pixel by pixel
    If KeyCode = vbKeyF Then
        cmdStart(Index).Left = cmdStart(Index).Left + 1
    ElseIf KeyCode = vbKeyA Then
        cmdStart(Index).Left = cmdStart(Index).Left - 1
    ElseIf KeyCode = vbKeyE Then
        cmdStart(Index).Top = cmdStart(Index).Top - 1
    ElseIf KeyCode = vbKeyC Then
        cmdStart(Index).Top = cmdStart(Index).Top + 1
    End If
End Sub

Private Sub cmdStart_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        ReleaseCapture
        SendMessage cmdStart(Index).hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
    End If
End Sub

Private Sub RedrawRects(showRgnRect As Boolean)

    If hRgn = 0 Then Exit Sub
    
    Dim hBrush As Long, x As Long, rgnRect() As RECT
    
    picGraph.BackColor = RGB(190, 190, 190)
    picGraph.Cls
    
    hBrush = CreateSolidBrush(vbWhite)
    FillRgn picGraph.hdc, hRgn, hBrush
    DeleteObject hBrush
    
    If showRgnRect Then
        rgnRect = ExtractRectangles(hRgn)
        hBrush = CreateSolidBrush(vbCyan)
        For x = 0 To UBound(rgnRect)
            FrameRect picGraph.hdc, rgnRect(x), hBrush
        Next
        DeleteObject hBrush

    End If
    
End Sub

Private Sub ClearPaths()

    Dim x As Long
    For x = 0 To lnPath.UBound
        lnPath(x).Visible = False
    Next
    For x = 0 To ln1st.UBound
        ln1st(x).Visible = False
    Next

End Sub

Private Sub pFinder_PathFound(ByVal isFinal As Boolean, Continue As Boolean)
' Optional Continue parameter: If set to false when this event is processed,
' then the path finding routine will cease finding any other paths.
' When the isFinal parameter is set to True, no other paths will be attempted


If pFinder.pathLength = 0 Then ' no path could be found when UBound is Zero
    
    txtPath(1).Text = "No possible path"
    txtPath(0).Text = "Is the A or Z target in an obstacle?"
    
Else
    
    Dim timerStop As Long
    
    pathCount = pathCount + 1
    
    If isFinal Then
        ' this is the final path
        timerStop = GetTickCount
        ' property returns a zero-bound, 2 dimensional array of X,Y coordinates of path points
        FinalPath = pFinder.PathPoints
        showPath True
        txtPath(1).Text = "Final (Black) length: " & FormatNumber(pFinder.pathLength, 2) & " (" & timerStop - timerStart & " ms)"
    
    ElseIf pathCount = 1 Then
        ' first path found, there may be more
        timerStop = GetTickCount
        If chkNoInitPath = 0 Then
            FinalPath = pFinder.PathPoints
            showPath False
        End If
        txtPath(0).Text = "Initial (Red) length: " & FormatNumber(pFinder.pathLength, 2) & " (" & timerStop - timerStart & " ms)"
    End If
    

End If

End Sub




Private Function ExtractRectangles(rgnMap As Long) As RECT()

' Homegrown function to extract the rectangle structure of a windows region
' Region rectangles can be extracted as a byte array & always follow a
' specific pattern: no rectangle ever has another rectangle that share an
' adjacent vertical edge (never side to side). A rectangle may have one,
' more than one or no rectangles vertically adjacent (top to bottom).

Dim rSize As Long, vRgnData() As Byte, vRect() As RECT
Dim y As Long, x As Long

    ' 1st get the buffer size needed to return rectangles info from this region
    rSize = GetRegionData(rgnMap, ByVal 0&, ByVal 0&)
    If rSize > 0 Then   ' success
        ' create the buffer & call function again to fill the buffer
        ReDim vRgnData(0 To rSize - 1) As Byte
        If rSize = GetRegionData(rgnMap, rSize, vRgnData(0)) Then     ' success
        
            ' Here are some tips for the structure returned
            ' Bytes 8-11 are the number of rectangles in the region
            ' Bytes 12-15 is structure size information -- not important for what we need
            ' Bytes 16-31 are the bounding rectangle's dimensions
            ' Bytes 32 to end of structure are the individual rectangle's dimensions
            ' The rectangle structure (RECT) is 16 bytes or Len(RECT)
        
            ' Let's retrieve the number of rectangles in the structure (b:8-11)
            CopyMemory rSize, vRgnData(8), ByVal 4&
            ReDim vRect(0 To rSize - 1)
            CopyMemory vRect(0), vRgnData(32), rSize * 16
        End If
    End If
    Erase vRgnData
    
    ' As long as we are here, we will reduce the number of rectangles up to 75%
    ' by combining those vertically adjacent rectangles that have the exact
    ' same width & left/right coordinates. They look like perfectly stacked bricks.
    
    ' This will reduce the total number of nodes we need to look at for path finding
    rSize = rSize - 1
    Do Until x >= rSize ' rsize will decrement so a For:Loop shouldn't be used
        y = x + 1       ' same reasoning here ^^
        If y < rSize Then
            Do
                ' see if the rectangles are vertically adjacent
                If vRect(x).Bottom = vRect(y).Top Then
                    ' now see if they are the same width
                    If vRect(y).Right = vRect(x).Right Then
                        ' and if they have the same left edge
                        If vRect(y).Left = vRect(x).Left Then
                            ' ok, let's merge the two
                            vRect(x).Bottom = vRect(y).Bottom
                            ' now shift the array one to the left
                            CopyMemory vRect(y), vRect(y + 1), 16 * (UBound(vRect) - y)
                            ' adjust the counters
                            y = y - 1
                            rSize = rSize - 1
                        End If
                    End If
                End If
                ' If row of rectangles will never be adjacent, exit this loop
                ' otherwise increment the counter
                If vRect(y).Top > vRect(x).Bottom Then Exit Do
                y = y + 1
            Loop Until y > rSize
        End If
        x = x + 1   ' increment the counter
    Loop
    ' almost impossible not to have merged some rectangles
    ' so instead of checking, simply redim the rectangle array
    ReDim Preserve vRect(0 To rSize)
    
ExtractRectangles = vRect

End Function

