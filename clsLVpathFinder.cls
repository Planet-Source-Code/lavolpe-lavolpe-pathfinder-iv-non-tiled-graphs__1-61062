VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLVpathFinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Not easy to follow unless you are very familiar with the A-Star (A*)
' path finding algorithm and how Window's regions are stored.
' Making it even more complicated is the sophisticated way I use Linked Lists
' and arrays to transform window region rectangles into horizontal edges to
' path points (anchors).

' My goal was to find paths in l00 ms or less. I have not quite accomplished that,
' but got real close (on average is less than 100 ms). That being said, this is
' the next step of something I started two years ago. There will be more, but this
' will be put to bed for awhile to allow me to think about ways to improve this project.
' Since no others I know of have attempted to path find using region rectangles,
' I get to make the rules & you get to bend or break them.

' I am always eager to talk with others that are interested in pursuing this unique
' logic in path finding on non-tiled graphs. If you want to tweak or improve this
' routine & are unsure of why I did what I did, simply email me with your questions.

' Until final tweaks are made. Some terminology that you may find confusing:
' Edge: the top edge of a rectangle. Unique & never duplicated
' Node: can refer to a rectangle or edge
' Anchor: a point in a path that the path MUST go thru
'   all paths have at least 2 anchors: starting & ending point


' Always declare your class like: Private myClass WithEvents As clsLVpathFinder
Public Event PathFound(ByVal isFinal As Boolean, ByRef Continue As Boolean)
' ^^ The only public event in this class. This event notifies you that a path
' has been found, whether or not it is the final path, and offers the option
' for you to abort continuing path finding, if desired

' very few APIs used
Private Declare Function PtInRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function InflateRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function GetRegionData Lib "gdi32" (ByVal hRgn As Long, ByVal dwCount As Long, lpRgnData As Any) As Long
Private Declare Function IntersectRect Lib "user32.dll" (ByRef lpDestRect As RECT, ByRef lpSrc1Rect As RECT, ByRef lpSrc2Rect As RECT) As Long

' Heavy use of UDTs to help keep arrays managable/readable
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type LinkData               ' individual links btwn region rectangles
    Ref As Long                     ' Pointer to Rectangle array item
    Edge As Long                    ' Pointer to related Edges array item
End Type
Private Type LinkReference          ' collection of LinkData elements
    k_Links As Long                 ' number of elements
    k_Neighbor() As LinkData        ' the collection
End Type

' Anchors are points in a path that cannot be removed from the path
Private Type AnchorData             ' Anchor Details
    Parent As Long
    x As Long                       ' X coordinate of point
    Edge As Long                    ' Ref to Edges() array which has the Y coord
    LowCost As Single               ' Min cost for this point to the path starting point
End Type
' SubAnchors are same as anchors but are unremovable points btwn two Anchors
Private Type SubAnchorData
    Count As Long                   ' Number of subanchors
    Anchor() As Long                ' Array of pointers into the Anchors() array
End Type

Private Type NodeReference          ' individual nodes
    Parent As Long                  ' node that generated this node
    Cost As Single                  ' Cost involved in exploring this node (filterable value)
    Anchor As Long                  ' Last anchor point for this node
    SubAnchor As Long               ' Ref to Subanchors() array, if any subanchors
End Type
Private Type PendingNodeReference   ' Pending cached path
    Count As Long                   ' Nr of nodes in the path
    Node() As NodeReference         ' Collection of the nodes
    Index As Long                   ' The last node index in the pending path
End Type
    
Private Type EdgeReference          ' individual Edges (horizontal only)
    x As Long                       ' left coordinate of the edge
    Z As Long                       ' right coordinate of the edge
    y As Long                       ' Y coordinate of the edge
    Count As Long                   ' Number of linked edges
    Edge() As Long                  ' Pointers to the links in the Edges array
End Type

' Enumerators only used in local functions/routines
Private Enum HeapActions
    heapPush = 0
    heapPop = 1
    heapModify = 2
End Enum

' As the items in some arrays can potentially be huge if we populated them in
' advance, several arrays are appended only when needed. This is slower overall
' but reduces the memory required to find extremely complicated paths on large graphs

Private rgnRect() As RECT           ' collection of rectangles in the map region
Private Links() As LinkReference    ' collection of rectangle links
Private Nodes() As NodeReference    ' collection of Nodes considered for paths
Private Edges() As EdgeReference    ' collection of Edge links
Private Anchors() As AnchorData     ' collection of path Anchor points
Private EdgeAnchors() As Long       ' cross-ref of Edges>Anchors
Private subAnchors() As SubAnchorData   ' collection of subAnchors

Private openList() As Long          ' binary heap of active nodes
Private pendingList() As Long       ' binary heap of pending nodes
Private nodesPending() As PendingNodeReference   ' Created as needed
' ^^ Unfortunately, window regions don't allow for simple path finding rules
'    Collection of nodes that need to be revisited
Private pendingIndex() As Long      ' list of reusable pending nodes indexes
' ^^ The nodesPending() array is a cached history of a path. This has potential
'    of being somewhat large & is often updated. To prevent rediming this array
'    which has subarrays and its sorted binary heap when items are removed from
'    the heap, we will simply keep track of those removed items & reuse them
'    as new nodes are placed on the heap (pendingList() array)

Private FinalPath() As Long         ' 2-dim array of key points in a path
Private FinalCost As Single         ' the cost/length of the final path

Private anchorPtLoopCount As Long       ' debug/testing use only

Public Property Get pathLength() As Single
    pathLength = FinalCost
End Property

Public Property Get PathPoints() As Long()
    If FinalCost = 0 Then ReDim FinalPath(0)
    PathPoints = FinalPath()
End Property

' ///////////////////// PATH ALGORITHM ROUTINES \\\\\\\\\\\\\\\\\\\\\

Public Sub FindPath(startPtX As Long, startPtY As Long, endPtX As Long, endPtY As Long, graphRgn As Long)
' startPtX & startPtY are the X,Y coordinates of the path's start position
' endPtX & endPtY are the X,Y coordinates of the path's end/terminal position
' graphRgn is a region containing the graph minus all obstacles (free space only)
'   ::use APIs like CombineRgn, CreateRectRgn, etc to create the region

anchorPtLoopCount = 0

Dim bContinue As Boolean
bContinue = True    ' flag to trigger Event at end of routine
' reset the cost before we start pathfinding
FinalCost = 0

' no region, no path
If graphRgn = 0 Then GoTo CleanUp

' extract region rectangles from passed region
ExtractRectangles graphRgn

If UBound(rgnRect) = 0 Then
    ' the entire graph is free space
    ReDim Nodes(0)
    ReDim Edges(1 To 2)
    ' ensure the points are on the graph
    If LocateTargets(startPtX, startPtY, endPtX, endPtY) = 0 Then GoTo CleanUp
    ' calculate the cost & set the points
    Nodes(0).Cost = Sqr(Abs(startPtX - endPtX) ^ 2 + Abs(startPtY - endPtY) ^ 2)
    SetFinalPath
    GoTo CleanUp
End If

' build edge and rectangle link arrays
BuildLinks
' resize the base node array & related Anchors array
ReDim Nodes(0 To UBound(Edges) + 2)
' ^^ Note: Nodes are ref'd as 1-bound. The 0-array item is used as a temporary node
'    common to a few of the routines in this project. No need to pass it to the function

' locate the targets on the graph. If any target is in an obstacle, abort.
Select Case LocateTargets(startPtX, startPtY, endPtX, endPtY)
    Case 0: ' one or both obstacles are off the graph, maybe in an obstacle
        GoTo CleanUp
    Case 2: ' both obstacles are in the same regional rectangle
        Nodes(0).Cost = Sqr(Abs(startPtX - endPtX) ^ 2 + Abs(startPtY - endPtY) ^ 2)
        SetFinalPath
        GoTo CleanUp
    Case Else
End Select

' A* (A Star) algorithm. There are many variations of this algorithm.
' A* is, in fact, a variation of Dijkstra's algorithm.
' Where A* requires a known "path traveled so-far" cost and an
'   estimated cost to end of path, I may not know the true value of
'   "path traveled so-far". The core logic of A* modified for this situation


ReDim Anchors(0 To 1)       ' reset path anchor points. Created as needed
ReDim subAnchors(0)         ' reset path subanchor points. Created as needed

ReDim EdgeAnchors(1 To UBound(Nodes) * 2)
' ^^ max number of possible anchors. Resets cross-reference table

' binary heap & node array flags
Dim nodeState() As Byte             ' state of any node
ReDim nodeState(1 To UBound(Edges))
ReDim openList(1 To 100)            ' binary heap of active nodes
Dim nrHeapItems As Long             ' number of items in the binary heap
Dim closedState As Byte             ' flag to indicate status of any node

ReDim nodesPending(1)               ' pending nodes to be revisited (as needed)
ReDim pendingList(1 To 100)         ' binary heap of pending nodes
Dim nrPending As Long               ' number of items in the binary heap
ReDim pendingIndex(0 To 1)          ' reusable nodesPending() array item Indexes

Dim nodeID As Long, pNode As Long   ' refrences into the Nodes Array
Dim adjNode As Long, pNodeIndex As Long ' and nodesPending() array
    
Dim bAbort As Boolean               ' flag to indicate continuance of path

' Add the starting node to the open list to be checked
' Node 1 will always be the starting location
' Node 2 will always be the ending location
nrHeapItems = 1
openList(1) = 1
' prime the pending nodes reusable index array
pendingIndex(0) = 1

With Anchors(0)         ' create two initial anchor points
    .Edge = 1           ' starting point
    .x = Edges(1).x
End With
With Anchors(1)         ' ending point
    .Edge = 2
    .x = Edges(2).x
End With

' set the value of this flag
closedState = 1

Do  ' double loop.
    ' one loop processes all current possible paths
    ' the 2nd loop adds any secondary paths onto the 1st loop after 1st loop is done

    Do ' continue until path found or no path can be found
        
        nodeState(openList(1)) = closedState 'update node as closed/explored/done
        ' this will be the parent node for any nodes it explores/opens
        pNode = openList(1)
    
        ' Get next open node from the binary heap
        openList(1) = openList(nrHeapItems)
        nrHeapItems = nrHeapItems - 1
        ' Need to adjust the heap now
        UpdateHeap heapPop, nrHeapItems
        
        ' get the 1st possible adjacent edge/node to our parent node
        adjNode = 0
        nodeID = GetAdjacents(pNode, adjNode)
        
        Do Until nodeID = 0
            ' call function to calculate any anchor/subanchor points and also
            ' determine the cost of moving to those points
            If CalcAnchorPt(nodeID, pNode) Then
            
                If nodeID = 2 Then  ' we hit the ending point
                    If SetFinalPath() Then
                        bContinue = True
                        If UBound(FinalPath, 2) = 1 Then
                            ' a straight line path test.
                            ' Since the fastest route btwn 2 points is a straight line,
                            ' if we have only 2 points in the path....
                            nrHeapItems = 0
                            nrPending = 0
                        Else
                            ' send user notice that a path was found
                            ' If user aborts pathfinding, we abort here
                            RaiseEvent PathFound(False, bContinue)
                            If Not bContinue Then GoTo CleanUp
                        End If
                    End If
                    ' no need to allow this node to continue
                    bAbort = True
                Else
                    ' not the ending node, but do we have a path yet?
                    If FinalCost Then
                        ' yep. Is this path longer? If so, abort it
                        bAbort = (Nodes(0).Cost > FinalCost)
                    End If
                End If
                If bAbort Then
                    ' we aborted, reset flag
                    bAbort = False
                Else
                
                    If nodeState(nodeID) = closedState Then
                        ' if this path hits another existing path, we want to push this
                        ' onto a pending path array to process later. We cannot process
                        ' it now 'cause we do not know where it will next turn or
                        ' terminate at and we can't modify any node's
                        ' parent properties without knowing that information in advance
                        UpdatePendingIndex heapPop, pNodeIndex
                        AddToNodesPending nodeID, pNode, pNodeIndex
                        nrPending = nrPending + 1
                        pendingList(nrPending) = pNodeIndex
                        UpdatePendingHeap heapPush, nrPending
                    Else
                        ' update the node's properties
                        Nodes(nodeID) = Nodes(0)
                        ' increment the heap count
                        nrHeapItems = nrHeapItems + 1
                        ' add the node to the the binary heap
                        openList(nrHeapItems) = nodeID
                        UpdateHeap heapPush, nrHeapItems
                        'Update the node state
                        nodeState(nodeID) = closedState

                    End If
                    
                End If '  done updating current node
            
            End If ' done testing for excessive path length
            
            ' get next adjacent to this parent node
            nodeID = GetAdjacents(pNode, adjNode)
            
        Loop ' adjacent nodes
        
    ' when the heap is empty, we no longer have candidates to explore, yet.
    Loop Until nrHeapItems = 0

    If nrPending Then
        ' process pending nodes. Since any edge can only have one parent,
        ' nodes from different paths that come to a common node would require
        ' us to make a decision which path would be that node's parent. But
        ' we don't have all the points in the new path & can't make that decision
        ' Therefore, we added that node to a pending list & will process it now
        
        ' instead of redim'ing the NodeState array, we will simply change the
        ' value of closedState. This allows us to reuse the array
        ' up to 255 times without clearing it out via a Redim each time
        If closedState = 255 Then
            ReDim nodeState(0 To UBound(Nodes))
            closedState = 0
        End If
        closedState = closedState + 1
        
        ' get the last item on the list
        Do
            pNodeIndex = pendingList(1)
            pendingList(1) = pendingList(nrPending)
            nrPending = nrPending - 1
            UpdatePendingHeap heapPop, nrPending
            
            ' we can run a check here first. While the node was pending,
            ' the final path length could have been shortened which could now
            ' nullify this pending path. If so, we don't need to finish it.
            If FinalCost > nodesPending(pNodeIndex).Node(0).Cost Then
            
                ' path length short enough to allow it to continue
                ' repopulate the Nodes() array with the cached path
                With nodesPending(pNodeIndex)
                    For nodeID = 0 To .Count - 2
                        ' redesignate nodes & parent relationships for this path
                        Nodes(.Node(nodeID).Parent) = .Node(nodeID + 1)
                        ' show the node state as closed
                        nodeState(.Node(nodeID).Parent) = closedState
                    Next
                    ' now bring on the node that was added to the pending list
                    Nodes(.Index) = .Node(0)
                    ' change it's state & add it as the 1st item on the heap
                    nodeState(.Index) = closedState
                    openList(1) = .Index
                    nrHeapItems = 1
                End With
                UpdatePendingIndex heapPush, pNodeIndex
                Exit Do
            Else
                UpdatePendingIndex heapPush, pNodeIndex
            End If
        Loop Until nrPending = 0
    
    End If

Loop Until nrHeapItems = 0

CleanUp:
' explicitly clear arrays
Erase nodesPending
Erase Nodes
Erase Edges
Erase Anchors
Erase subAnchors
Erase EdgeAnchors
Erase openList
Erase pendingList
Erase pendingIndex
Erase nodeState

' send final event to user
If bContinue Then RaiseEvent PathFound(True, False)

Debug.Print "total calcAnchorPt loops til final path: "; anchorPtLoopCount

End Sub

Private Function GetAdjacents(nodeID As Long, nextIndex As Long) As Long
' function returns 1st adjacent edge starting with passd "nextIndex"
' for the target node/edge (NodeID)

' This function filters possible path segments to prevent backtracking
' and wasted rerouting

Dim x As Long           ' generic loop variable
Dim rtnNode As Long     ' return value
Dim adjNode As Long     ' generic loop variable
Dim bAllow As Boolean   ' filter flag
Dim pDirBase As Byte    ' parent to current node direction
Dim tDirection As Byte  ' current node to target direction
Dim pDirection As Byte  ' parent to current node direction

With Edges(nodeID)

    ' determine the primary direction between the parent node & the current node
    Select Case .y - Edges(Nodes(nodeID).Parent).y
    Case Is < 0: pDirBase = 2 ' northerly route
    Case Is > 0: pDirBase = 1 ' southerly direction
    End Select  ' otherwise moving horizonally
    
    For x = nextIndex To .Count - 1
        
        ' prevent backtracking/rerouting to save processing time
        ' 1st check. Is next node the current node's parent?
        If Nodes(nodeID).Parent <> .Edge(x) Then
        
            If nodeID = 1 Or .Edge(x) = 2 Then
                ' if we are on our very 1st node or the last node; then allow connections
                bAllow = True
            Else
                ' not the very first node; do some more checks
                pDirection = pDirBase
                ' determine primary direction between current node & next node
                Select Case Edges(.Edge(x)).y - .y
                Case Is < 0: tDirection = 1 ' northerly route
                Case Is > 0: tDirection = 2 ' southerly direction
                End Select  ' otherwise moving horizonally
                '^^ notice tDirection is opposite value of pDirection if same direction.
                '   this is done so we can simply use an AND operand in the final check
                
                ' 2nd check. Entering/Exiting a bridge
                If pDirection = 0 Or tDirection = 0 Then
                    ' Don't allow bridge to be entered/exited if bridge is also
                    ' directly adjacent to the parent node/edge
                    For adjNode = 0 To Edges(Nodes(nodeID).Parent).Count - 1
                        If Edges(Nodes(nodeID).Parent).Edge(adjNode) = .Edge(x) Then
                            pDirection = 3
                            tDirection = pDirection 'flag prevents bridge connections
                            Exit For
                        End If
                    Next
                End If
                
                ' 3rd Check. Only allow connection if not wasted reroute
                bAllow = ((pDirection And tDirection) = 0)
                
                If bAllow Then
                    ' 4th check proof positive not backtracking
                    ' start with grandfather since parent was checked at top
                    adjNode = Nodes(Nodes(nodeID).Parent).Parent
                    Do Until adjNode = 1
                        If .Edge(x) = adjNode Then
                            bAllow = False
                            Exit Do
                        End If
                        adjNode = Nodes(adjNode).Parent
                    Loop
                End If
                tDirection = 0  ' reset
            End If
            
            ' pass the edge reference if allowable
            If bAllow Then
                rtnNode = .Edge(x)  ' set return value
                nextIndex = x + 1   ' set next index value
                Exit For
            End If
            
        End If  ' test if next node is also the parent node
    
    Next    ' check next adjacent node if necessary
    
End With

GetAdjacents = rtnNode

End Function

Private Sub AddToNodesPending(nodeIndex As Long, newParent As Long, pIndex As Long)

Dim nrNodes As Long
Dim pNode As Long

' resize the array as needed; giving a small buffer to prevent resizing more often
If pIndex > UBound(nodesPending) Then
    ReDim Preserve nodesPending(0 To UBound(nodesPending) + 20)
End If

' count the number of edge segments in this current path
pNode = newParent
Do Until pNode = 1
    nrNodes = nrNodes + 1
    pNode = Nodes(pNode).Parent
Loop

' resize the array Path & Anchor elements appropriately
ReDim nodesPending(pIndex).Node(0 To nrNodes)

' populate the array Path & Anchor elements
nrNodes = 1
pNode = newParent
Do Until pNode = 1
    nodesPending(pIndex).Node(nrNodes) = Nodes(pNode)
    nrNodes = nrNodes + 1
    pNode = Nodes(pNode).Parent
Loop
' fill in the core elements
With nodesPending(pIndex)
    .Count = nrNodes
    .Index = nodeIndex
    .Node(0) = Nodes(0)
End With

End Sub

Private Function CalcHueristics(fromAnchor As Long, fromNode As Long) As Single

' Still a work in progress, playing with different heuristic calculations
' This is the MOST important routine. The closer the first path found is to
' the final path, dramatically reduces the overall time needed to find the best path

Dim lngDiff As Long
Dim rtnVal As Single, midX As Single
Dim basePt As POINTAPI, endPt As POINTAPI
Dim bIsNotValid As Boolean, bIsLateral As Boolean

With Anchors(Nodes(0).Anchor)

    If fromNode = 2 Then
        ' moving to the final node (path ending point). Get exact distance
        rtnVal = Sqr(Abs(.x - Edges(2).x) ^ 2 + Abs(Edges(.Edge).y - Edges(2).y) ^ 2)
    Else
        ' moving from the last anchor point to some other regional edge
        ' This is the guess in distance. But the key is NOT to overestimate
        ' the distance to a value that would be longer than the actual distance.
        ' The trick is to be as close to exact as possible.
        ' Since we don't know the next turn point or even the next rectangle that
        ' will be visited, we have no way of knowing the exact distance!
        
        ' See if the edge being checked is physically between the last anchor
        ' and the ending point
        If Edges(.Edge).y < Edges(2).y Then
            If Edges(fromNode).y < Edges(.Edge).y Then bIsNotValid = True
        Else
            If Edges(fromNode).y > Edges(2).y Then bIsNotValid = True
        End If
        If bIsNotValid Then
            ' if the above is the case, we handle guesswork differently
            If Edges(2).x < Edges(fromNode).x Then
                midX = Edges(fromNode).x
            ElseIf Edges(2).x > Edges(fromNode).Z Then
                midX = Edges(fromNode).Z
            Else
                midX = Edges(2).x
            End If
            rtnVal = Sqr(Abs(midX - Edges(2).x) ^ 2 + Abs(Edges(fromNode).y - Edges(2).y) ^ 2)
            rtnVal = rtnVal + Abs(Edges(fromNode).y - Edges(.Edge).y)
            
        Else
            ' the visited edge is between the last known anchor & the ending point
            ' Draw a diagonal line thru that edge & get an intersection point closest
            ' to that edge. Use that point for measurement
            basePt.x = .x
            basePt.y = Edges(.Edge).y
            endPt.x = Edges(2).x
            endPt.y = Edges(2).y
            midX = LineIntersectCoord(basePt, endPt, Edges(fromNode).y, bIsLateral)
            If bIsLateral Then midX = .x
            If midX < Edges(fromNode).x Then
                midX = Edges(fromNode).x
            ElseIf midX > Edges(fromNode).Z Then
                midX = Edges(fromNode).Z
            End If
            ' measure from that point to last known anchor
            lngDiff = Abs(Edges(fromNode).y - Edges(.Edge).y)
            rtnVal = Sqr(Abs(.x - midX) ^ 2 + lngDiff ^ 2)
            ' measure from that pont to the end point
            lngDiff = Abs(Edges(fromNode).y - Edges(2).y)
            rtnVal = rtnVal + Sqr(Abs(Edges(2).x - midX) ^ 2 + lngDiff ^ 2)
        End If
    End If
    
    ' carry over the cost from the parent node (accumulative), sort of.
    rtnVal = rtnVal + .LowCost
        
End With

CalcHueristics = rtnVal

End Function

Private Sub UpdateHeap(Action As HeapActions, heapOffset As Long)
' BINARY HEAP ROUTINE for active nodes

' Don't mess with the binary heap routines.
' Binary heaps basically move every other item in an array until the sort order
' is done; but it is a bit more complicated than that. If unsure what binary heaps
' are, suggest pulling down a few web pages to get familiar with it. If you foul up
' the heap routines: at best a failed sorted list will occur,
' at worst: infinite loop

Dim heapPtr1 As Long
Dim heapPtr2 As Long
Dim SwapTemp As Long

Select Case Action
Case heapPop
    ' We want to keep the Node with lowest F cost at top of heap
    heapPtr2 = 1
    Do
        heapPtr1 = heapPtr2
        If heapPtr1 * 2 + 1 <= heapOffset Then
        ' see if at least 2 more items on heap exist
            If Nodes(openList(heapPtr1)).Cost >= Nodes(openList(heapPtr1 * 2)).Cost Then heapPtr2 = heapPtr1 * 2
            If Nodes(openList(heapPtr2)).Cost >= Nodes(openList(heapPtr1 * 2 + 1)).Cost Then heapPtr2 = heapPtr1 * 2 + 1
        
        Else    ' if only one more item on heap exists, do that one
            If heapPtr1 * 2 <= heapOffset Then
                If Nodes(openList(heapPtr1)).Cost >= Nodes(openList(heapPtr1 * 2)).Cost Then heapPtr2 = heapPtr1 * 2
            End If
        End If
        
        If heapPtr1 <> heapPtr2 Then
            ' swap the pointers
            SwapTemp = openList(heapPtr1)
            openList(heapPtr1) = openList(heapPtr2)
            openList(heapPtr2) = SwapTemp
        Else
            Exit Do
        End If
    Loop ' until done
Case heapPush
    ' add a node to the heap
    If heapOffset + 1 > UBound(openList) Then
        ReDim Preserve openList(1 To UBound(openList) + 100)
    End If
    heapPtr1 = heapOffset
    Do While heapPtr1 <> 1
        If Nodes(openList(heapPtr1)).Cost <= Nodes(openList(heapPtr1 / 2)).Cost Then
            SwapTemp = openList(heapPtr1 / 2)
            openList(heapPtr1 / 2) = openList(heapPtr1)
            openList(heapPtr1) = SwapTemp
            heapPtr1 = heapPtr1 / 2
        Else
            Exit Do
        End If
    Loop
   
End Select
End Sub

Private Sub UpdatePendingHeap(Action As HeapActions, heapOffset As Long)
' BINARY HEAP ROUTINE for pending nodes

' Don't mess with the binary heap routines.
' Binary heaps basically move every other item in an array until the sort order
' is done; but it is a bit more complicated than that. If unsure what binary heaps
' are, suggest pulling down a few web pages to get familiar with it. If you foul up
' the heap routines: at best a failed sorted list will occur,
' at worst: infinite loop

Dim heapPtr1 As Long
Dim heapPtr2 As Long
Dim SwapTemp As Long

Select Case Action
Case heapPop
    ' We want to keep the Node with lowest F cost at top of heap
    heapPtr2 = 1
    Do
        heapPtr1 = heapPtr2
        If heapPtr1 * 2 + 1 <= heapOffset Then
        ' see if at least 2 more items on heap exist
            If nodesPending(pendingList(heapPtr1)).Node(0).Cost >= nodesPending(pendingList(heapPtr1 * 2)).Node(0).Cost Then heapPtr2 = heapPtr1 * 2
            If nodesPending(pendingList(heapPtr2)).Node(0).Cost >= nodesPending(pendingList(heapPtr1 * 2 + 1)).Node(0).Cost Then heapPtr2 = heapPtr1 * 2 + 1
        
        Else    ' if only one more item on heap exists, do that one
            If heapPtr1 * 2 <= heapOffset Then
                If nodesPending(pendingList(heapPtr1)).Node(0).Cost >= nodesPending(pendingList(heapPtr1 * 2)).Node(0).Cost Then heapPtr2 = heapPtr1 * 2
            End If
        End If
        
        If heapPtr1 <> heapPtr2 Then
            ' swap the pointers
            SwapTemp = pendingList(heapPtr1)
            pendingList(heapPtr1) = pendingList(heapPtr2)
            pendingList(heapPtr2) = SwapTemp
        Else
            Exit Do
        End If
    Loop ' until done
Case heapPush
    ' add a node to the heap
    If heapOffset + 1 > UBound(pendingList) Then
        ReDim Preserve pendingList(1 To UBound(pendingList) + 100)
    End If
    heapPtr1 = heapOffset
    Do While heapPtr1 <> 1
        If nodesPending(pendingList(heapPtr1)).Node(0).Cost <= nodesPending(pendingList(heapPtr1 / 2)).Node(0).Cost Then
            SwapTemp = pendingList(heapPtr1 / 2)
            pendingList(heapPtr1 / 2) = pendingList(heapPtr1)
            pendingList(heapPtr1) = SwapTemp
            heapPtr1 = heapPtr1 / 2
        Else
            Exit Do
        End If
    Loop
   
End Select
End Sub

Private Sub UpdatePendingIndex(Action As HeapActions, Index As Long)
' Routine keeps track of which pendingNodes() array items can be reused

If Action = heapPop Then ' need next reusable index
    
    If pendingIndex(0) = 0 Then
        ' no resuable indexes available, therefore; set to UBound+1
        Index = UBound(nodesPending) + 1
    Else
        ' (0) is pointer to array item to be used as the Index
        Index = pendingIndex(pendingIndex(0))
        pendingIndex(0) = pendingIndex(0) - 1
    End If

Else                    ' adding a reusable index
    
    ' this portion of the nodesPending() array can be large; erase it
    Erase nodesPending(Index).Node
    If pendingIndex(0) = UBound(pendingIndex) Then
        ' resize the array giving a small buffer to prevent resizing more often
        ReDim Preserve pendingIndex(0 To UBound(pendingIndex) + 10)
    End If
    ' (0) is pointer to next Index to use
    pendingIndex(0) = pendingIndex(0) + 1
    ' update the array with the Index
    pendingIndex(pendingIndex(0)) = Index
End If

End Sub

Private Function LineIntersectCoord(PointA As POINTAPI, PointB As POINTAPI, y As Long, isBridged As Boolean) As Single
' Basic line intersection routine.

' PointA & PointB comprise the X,Y coords of one of the lines
' Since one line is always 180*, only Y need to be passed for the second line
' Additionally, standard intersection routines test to see if the collision is
' on both lines; not needed here 'cause I want to predict where the collision
' would occur on the passed PointA/B line.

' Modified to tweak & exclude unneeded calculations.
' One line will always be 180* horizontal, so there is no need to calculate
' the slope of that line or the Y coordinate of intersection

Dim x As Single
If PointA.x = PointB.x Then ' vertically parallel points
    
    ' therefore a 90* & 180* intersection
    ' simply return the vertical X coordinate
    x = PointA.x
    isBridged = False

ElseIf PointA.y = PointB.y Then

    'horizontally parallel points
    isBridged = True

Else
    ' calculate intersection X coordinate
    Dim M1 As Single, B1 As Single
    M1 = (PointB.y - PointA.y) / (PointB.x - PointA.x)
    B1 = PointA.y - M1 * PointA.x
    x = (y - B1) / M1
    isBridged = False
    
End If

' Return where intersection would occur
LineIntersectCoord = x

End Function

Private Function CalcAnchorPt(nodeID As Long, parentNode As Long) As Boolean
' A core routine, 2nd most imporant routine. Detemines if an edge can be reached
' directly from a given point without obstruction.  If obstructions exists, then the
' routine creates anchor points / turn points for the path. This routine is the
' most time consuming portion of this class. We take significant steps to prevent
' running this routine if possible. Some steps/filters were applied before this
' routine was called and others are applied during this routine

Dim x As Single             ' generic counter
Dim pNode As Long           ' node ref used in a Do Loop
Dim bBridged As Boolean     ' booleans to determine anchor type/aborts
Dim bDone As Boolean        ' flag to terminate the double loop below
Dim bAbort As Boolean       ' flag to terminate this routine

Dim midAnchor As Long       ' possible anchor point
Dim basePt As POINTAPI      ' used to determine if path can pass thru
Dim targetPt As POINTAPI    '   an edge without obstruction

Dim bHasSubAnchor As Boolean    ' used to identify if subAnchors() array
Dim newAnchor As Long           ' needs to be appended to

Dim anchorCost As Single        ' return value of calculating distances

' Following Do:Loop is a filter to reduce number of times this routine is run

' The logic here is... If a node is being processed and some point down
' its path has been changed (differnt parent assigned), then this node can be
' aborted. This is because whichever node got to that point faster changed the
' parent and is still being processed/considered in the path,
' it wasn't aborted & this one can be.

' Check each anchor and subanchor to see if it changed parent
' Quick since any path will not have many anchor points in it

pNode = parentNode
newAnchor = Nodes(parentNode).Anchor
Do Until newAnchor = 0
    If Nodes(pNode).Anchor <> newAnchor Then
        If Anchors(newAnchor).Parent <> Nodes(pNode).Anchor Then
            bAbort = True
            Exit Do
        End If
        newAnchor = Nodes(pNode).Anchor
    End If
    If Nodes(pNode).SubAnchor Then
        For midAnchor = subAnchors(Nodes(pNode).SubAnchor).Count - 1 To 0 Step -1
            If Anchors(newAnchor).Parent <> subAnchors(Nodes(pNode).SubAnchor).Anchor(midAnchor) Then
                bAbort = True
                Exit Do
            End If
            newAnchor = subAnchors(Nodes(pNode).SubAnchor).Anchor(midAnchor)
        Next
    End If
    pNode = Nodes(pNode).Parent
Loop
If bAbort Then Exit Function

' set up the temp Node. Prevents undoing actions if we abort later
Nodes(0) = Nodes(nodeID)
Nodes(nodeID).Parent = parentNode   ' incoming node generally won't have a parent yet
Nodes(0).SubAnchor = 0              ' ensure no subanchors are referenced
' get the parent nodes last turn point, this is our goal for path segment
Nodes(0).Anchor = Nodes(parentNode).Anchor

' set the base point of our diagonal line to the parent's last turn point
basePt.x = Anchors(Nodes(0).Anchor).x
basePt.y = Edges(Anchors(Nodes(0).Anchor).Edge).y
    
' see if we can get diagonal from last turn point to our new edge
' without obstruction.  If not, find the nearest new turn point to
' the base point

Do
    ' set the other end point of the diagonal line to the far edge of the target edge
    ' Note: Setting it to any point on the edge will work with this algorithm
    targetPt.x = Edges(nodeID).Z
    targetPt.y = Edges(nodeID).y
    ' identify 1st node/edge to be checked
    pNode = parentNode
    midAnchor = 0       ' reset
    
    ' continue checking up the path until we get to our basePoint's edge/node
    Do Until pNode = Anchors(Nodes(0).Anchor).Edge
        anchorPtLoopCount = anchorPtLoopCount + 1
        ' see if this edge intersects the diagonal line
        x = LineIntersectCoord(basePt, targetPt, Edges(pNode).y, bBridged)
        
        ' horizontal movements are handled uniquely
        If bBridged Then
            If midAnchor = 0 Then
                ' bridge to bridge to bridge
                If Edges(pNode).x > Edges(nodeID).Z Then
                    ' moving east>west
                    anchorCost = SetAnchor(pNode, Edges(pNode).x, Nodes(0).Anchor, newAnchor, nodeID)
                Else ' west>east
                    anchorCost = SetAnchor(pNode, Edges(pNode).Z, Nodes(0).Anchor, newAnchor, nodeID)
                End If
            Else ' at least one bridge involved; may be two or more
                If Anchors(Nodes(0).Anchor).x > Edges(midAnchor).Z Then
                    anchorCost = SetAnchor(midAnchor, Edges(midAnchor).Z, Nodes(0).Anchor, newAnchor, nodeID)
                Else
                    anchorCost = SetAnchor(midAnchor, Edges(midAnchor).x, Nodes(0).Anchor, newAnchor, nodeID)
                End If
            End If
            
            If anchorCost Then  ' anchor was set or is now shorter than before
                ' set up subanchors if needed
                If bHasSubAnchor Then
                    If Nodes(0).SubAnchor = 0 Then
                        ReDim Preserve subAnchors(0 To UBound(subAnchors) + 1)
                        Nodes(0).SubAnchor = UBound(subAnchors)
                    End If
                    With subAnchors(Nodes(0).SubAnchor)
                        ReDim Preserve .Anchor(0 To .Count)
                        .Anchor(.Count) = Nodes(0).Anchor
                        .Count = .Count + 1
                    End With
                Else
                    bHasSubAnchor = True
                End If
                ' cache node's latest anchor point
                Nodes(0).Anchor = newAnchor
            Else
                ' we will not continue. This path segment is longer than one
                ' that has previously gone thru the anchor
                
                If Nodes(0).SubAnchor Then ReDim Preserve subAnchors(0 To Nodes(0).SubAnchor - 1)
                ' ^^optional. But keep arrays lengths at a minimum if possible
                bAbort = True
            End If
            midAnchor = 0   ' prevents secondary loop from triggering
            Exit Do
        End If
        ' obstacle at this point. Change angle of diagonal line to
        ' go around this obstacle & try again
        If x < Edges(pNode).x Then
            targetPt.x = Edges(pNode).x
            targetPt.y = Edges(pNode).y
            midAnchor = pNode
        ElseIf x > Edges(pNode).Z Then
            targetPt.x = Edges(pNode).Z
            targetPt.y = Edges(pNode).y
            midAnchor = pNode
        End If
        ' move to the previous parent
        pNode = Nodes(pNode).Parent
    Loop
    
    ' secondary loop....
    ' The 1st loop identified the anchor closest to the previous anchor that
    ' will now be included in the path. However, we may still have other
    ' anchors that need to be calculated. This loop will check from the
    ' new anchor to our current node. If other obstructions occur, then
    ' we will continue the dual loop until no more anchors are needed.
    
    If midAnchor Then   ' obstacle blocking a straight path
        ' See if we can get to the new anchor without obstruction
        pNode = nodeID
        Do Until pNode = midAnchor
            anchorPtLoopCount = anchorPtLoopCount + 1
            x = LineIntersectCoord(basePt, targetPt, Edges(pNode).y, bBridged)
            
            If x < Edges(pNode).x Or x > Edges(pNode).Z Or bBridged = True Then
                ' obstacle cannot be bypassed
                anchorCost = SetAnchor(midAnchor, targetPt.x, Nodes(0).Anchor, newAnchor, nodeID)
                If anchorCost Then  ' anchor was set or is now shorter than before
                    ' set up subAnchors if needed
                    If bHasSubAnchor Then
                        If Nodes(0).SubAnchor = 0 Then
                            ReDim Preserve subAnchors(0 To UBound(subAnchors) + 1)
                            Nodes(0).SubAnchor = UBound(subAnchors)
                        End If
                        With subAnchors(Nodes(0).SubAnchor)
                            ReDim Preserve .Anchor(0 To .Count)
                            .Anchor(.Count) = Nodes(0).Anchor
                            .Count = .Count + 1
                        End With
                    Else
                        bHasSubAnchor = True
                    End If
                    ' cache latest anchor & adjust the diagonal line again
                    Nodes(0).Anchor = newAnchor
                    basePt = targetPt
                Else
                    ' we will not continue. This path segment is longer than one
                    ' that has previously gone thru the anchor
                    If Nodes(0).SubAnchor Then ReDim Preserve subAnchors(0 To Nodes(0).SubAnchor - 1)
                    ' ^^optional. But keep arrays lengths at a minimum if possible
                    midAnchor = pNode ' prevent 1st loop from triggering
                    bAbort = True
                End If
                Exit Do
            End If
            ' move to previous parent
            pNode = Nodes(pNode).Parent
        Loop
        ' if we got to the turn point, no need to continue the primary loop
        bDone = (midAnchor = pNode)
    Else
        ' we have an unobstructed possible path from last anchor to our edge
        bDone = True
    End If

    ' because there can be multiple turn points needed to move to our
    ' target edge, we may need to repeat as needed
Loop Until bDone


' replace the node's original parent reference
Nodes(nodeID).Parent = Nodes(0).Parent

If Not bAbort Then
    ' finalize changes
    With Nodes(0)
        ' set the parent node
        .Parent = parentNode
        ' calculate estimated cost to get to this edge
        .Cost = CalcHueristics(.Anchor, nodeID)
    End With

    CalcAnchorPt = True
    
End If
    
End Function

Private Function SetAnchor(aEdge As Long, aPt As Long, aParent As Long, newAnchor As Long, curNode As Long) As Single

' Function simply compares the new anchor cost against any pre-existing costs,
' returns that cost and a reference to an anchor to be used for the path

Dim tCost As Single     ' cost between anchor & starting point on existing path
Dim pAnchor As Long     ' return value for anchor to use
Dim anchorRef As Long   ' cross-ref value for EdgeAnchors() array

' get the cross-reference value
If aPt = Edges(aEdge).x Then
    anchorRef = aEdge * 2 - 1
Else
    anchorRef = aEdge * 2
End If

' calculate true distance btwn this new anchor & its parent anchor
tCost = Sqr(Abs(Anchors(aParent).x - aPt) ^ 2 + Abs(Edges(Anchors(aParent).Edge).y - Edges(aEdge).y) ^ 2) + Anchors(aParent).LowCost


If EdgeAnchors(anchorRef) Then

    ' anchor previously set, is the new cost longer, shorter or same?
    Select Case Anchors(EdgeAnchors(anchorRef)).LowCost - tCost
    
    Case Is > 0 ' shorter cost. Update the Anchors() array
        pAnchor = EdgeAnchors(anchorRef)
        ' remove this anchor from its current parent
        ' set anchor's new parent & cost
        Anchors(pAnchor).Parent = aParent
        Anchors(pAnchor).LowCost = tCost
        
    Case Is < 0 ' new path is longer
        Exit Function
        
    Case Else   ' same length; let it continue
        pAnchor = EdgeAnchors(anchorRef)
    End Select

Else
    
    ' no previous anchor point. Create one here
    pAnchor = UBound(Anchors) + 1
    EdgeAnchors(anchorRef) = pAnchor        ' update cross-reference array
    ReDim Preserve Anchors(0 To pAnchor)    ' append new anchor to Anchors() array
    With Anchors(pAnchor)
        .Parent = aParent
        .LowCost = tCost
        .Edge = aEdge
        .x = aPt
    End With
    
End If

' return the cost of the new anchor & its Index
SetAnchor = tCost
newAnchor = pAnchor

End Function

Private Function SetFinalPath() As Boolean

' function caches the best path as a 2-dimensional array of X,Y coordinates

If FinalCost Then
    ' previous path exists... If this one is longer, abort
    If Nodes(0).Cost >= FinalCost Then Exit Function
End If
FinalCost = Nodes(0).Cost  ' set the cost of the final path

Dim lastAnchor As Long
Dim nrAnchors As Long

' cache the path

' 1st. Count how many anchor points are in the path
lastAnchor = Nodes(0).Anchor
nrAnchors = 1
Do Until lastAnchor = 0
    nrAnchors = nrAnchors + 1
    lastAnchor = Anchors(lastAnchor).Parent
Loop
ReDim FinalPath(0 To 1, 0 To nrAnchors)


' set first point (starting point)
FinalPath(0, nrAnchors) = Edges(2).x
FinalPath(1, nrAnchors) = Edges(2).y
    
' fill in points between first & last point (in reverse order)
lastAnchor = Nodes(0).Anchor
Do Until lastAnchor = 0
    nrAnchors = nrAnchors - 1
    FinalPath(0, nrAnchors) = Anchors(lastAnchor).x
    FinalPath(1, nrAnchors) = Edges(Anchors(lastAnchor).Edge).y
    lastAnchor = Anchors(lastAnchor).Parent
Loop
' set first point (starting point), if needed
FinalPath(0, 0) = Edges(1).x
FinalPath(1, 0) = Edges(1).y

SetFinalPath = True

End Function


' /////////////////// GRAPH SETUP ROUTINES \\\\\\\\\\\\\\\\\\\\\\\\\\

Private Function ExtractRectangles(rgnMap As Long) As Long

' Homegrown function to extract the rectangle structure of a windows region
' Region rectangles can be extracted as a byte array & always follow a
' specific pattern: no rectangle ever has another rectangle that share an
' adjacent vertical edge (never side to side). A rectangle may have one,
' more than one or no rectangles vertically adjacent (top to bottom).

Dim rSize As Long, vRgnData() As Byte
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
            ReDim rgnRect(0 To rSize - 1)
            CopyMemory rgnRect(0), vRgnData(32), rSize * 16
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
                If rgnRect(x).Bottom = rgnRect(y).Top Then
                    ' now see if they are the same width
                    If rgnRect(y).Right = rgnRect(x).Right Then
                        ' and if they have the same left edge
                        If rgnRect(y).Left = rgnRect(x).Left Then
                            ' ok, let's merge the two
                            rgnRect(x).Bottom = rgnRect(y).Bottom
                            ' now shift the array one to the left
                            CopyMemory rgnRect(y), rgnRect(y + 1), 16 * (UBound(rgnRect) - y)
                            ' adjust the counters
                            y = y - 1
                            rSize = rSize - 1
                        End If
                    End If
                End If
                ' If row of rectangles will never be adjacent, exit this loop
                ' otherwise increment the counter
                If rgnRect(y).Top > rgnRect(x).Bottom Then Exit Do
                y = y + 1
            Loop Until y > rSize
        End If
        x = x + 1   ' increment the counter
    Loop
    ' almost impossible not to have merged some rectangles
    ' so instead of checking, simply redim the rectangle array
    ReDim Preserve rgnRect(0 To rSize)
    
ExtractRectangles = rSize

End Function

Private Sub BuildLinks()
' Function builds the link tables.
'   -- Links btwn Rectangles & links btwn Edges
' Rectangles are linked to adjacent rectangles and also
' Edges are linked to adjacent edges.
' Unlike rectangles which will never have horizontal links; edges can.

Dim lat1 As Long, lat2 As Long
Dim x As Long
Dim Neighbor As Long
Dim nextEdge As Long

' Creating the linked list tree
ReDim Links(0 To UBound(rgnRect))
' start with this many Edges; change as needed: may increase but mostly decreases
ReDim Edges(1 To UBound(rgnRect) * 2)

' Edge #1 is always the starting point in the path
' Edge #2 is always the ending point in the path
nextEdge = 3 ' start appending array with this index

For x = 0 To UBound(rgnRect) - 1
    
    ' get vertical neighbor, must have a common edge
    ' in other words, this Rect & adjacent must share some part of their
    ' bottom & top edges respectively
    Neighbor = GetRectNeighbor(x, x + 1, lat1, lat2)
    
    Do Until Neighbor < 0
        
        ' we have a vertically adjacent rectangle to this one
        With Links(x)   ' build link joining current one to the adjacent
            ReDim Preserve .k_Neighbor(0 To .k_Links)
            .k_Neighbor(.k_Links).Ref = Neighbor
            .k_Neighbor(.k_Links).Edge = nextEdge   ' identify the common edge
            .k_Links = .k_Links + 1
        End With
        With Links(Neighbor)    ' link the other side (bidirectional link)
            ReDim Preserve .k_Neighbor(0 To .k_Links)
            .k_Neighbor(.k_Links).Ref = x
            .k_Neighbor(.k_Links).Edge = nextEdge   ' identify the same edge
            .k_Links = .k_Links + 1
        End With
        ' now build the common edge
        With Edges(nextEdge)
            .x = lat1
            .Z = lat2 - 1
            .y = rgnRect(Neighbor).Top
        End With
        
        ' call function to join this common edge to any other previous edges
        ' that may have been created from previous links
        CreateBridgesVertical x, nextEdge
        
        ' increment the edge count & check for out of bounds
        nextEdge = nextEdge + 1
        If nextEdge + 10 > UBound(Edges) Then ReDim Preserve Edges(1 To UBound(Edges) + 100)
        
        ' Get the next vertical adjacent neighbor, if any
        Neighbor = GetRectNeighbor(x, Neighbor + 1, lat1, lat2)
    
    Loop
        
    ' after the rectangle has been checked vertically, we will
    ' attempt to add bridges immediately above & below the rectangle
    CreateBridgesTop x
    CreateBridgesBottom x
    
Next

' process the very last rectangle for bridges
CreateBridgesBottom x
' resize the array now
ReDim Preserve Edges(1 To nextEdge - 1)

End Sub

Private Sub CreateBridgesVertical(rIndex As Long, edgeIndex As Long)

' A vertical edge link can only exist if moving between the edges at any angle
' can be accomplished without moving through an obstacle

' Edge links are built in North to South order. Therefore, we only need to
' check for links between passed rectangle (rIndex) and those above it

Dim x As Long
For x = 0 To Links(rIndex).k_Links - 1
    ' find all adjacent rectangle adjacents above this one
    If rgnRect(Links(rIndex).k_Neighbor(x).Ref).Bottom = rgnRect(rIndex).Top Then
        ' build the edge links (bidirectional)
        With Edges(Links(rIndex).k_Neighbor(x).Edge)
            ReDim Preserve .Edge(0 To .Count)
            .Edge(.Count) = edgeIndex
            .Count = .Count + 1
        End With
        With Edges(edgeIndex)
            ReDim Preserve .Edge(0 To .Count)
            .Edge(.Count) = Links(rIndex).k_Neighbor(x).Edge
            .Count = .Count + 1
        End With
    End If
Next
End Sub

Private Sub CreateBridgesTop(rIndex As Long)

' Exercise in walking the Links Tree
' Build bridges between rectangles horizontally across the top of an obstacle
' that separates 2 rectangles. Create a bridge to the nearest right rectangle
' only if both rectangles share a common regional rectangle above them.

' Similar to Vertical links, this routine only moves left to right.

Dim x As Long, y As Long, V As Long, Adjacent As Long

For x = 0 To Links(rIndex).k_Links - 1
    ' determine which link may span across 2 rectangles with same Top edge value
    If rgnRect(Links(rIndex).k_Neighbor(x).Ref).Bottom = rgnRect(rIndex).Top Then
        Adjacent = -1
        ' this rectangle may allow a bridge.
        ' See if any of its links are parallel to our target
        With Links(rIndex).k_Neighbor(x)
            For y = 0 To Links(.Ref).k_Links - 1
                If rgnRect(Links(.Ref).k_Neighbor(y).Ref).Top = rgnRect(rIndex).Top Then
                    ' don't let our target be included here
                    If Links(.Ref).k_Neighbor(y).Ref <> rIndex Then
                        ' only look right >>>>
                        If rgnRect(Links(.Ref).k_Neighbor(y).Ref).Left > rgnRect(rIndex).Right Then
                            ' ok, now see if it is physically closer than any previous horizonal neighbor
                            If Adjacent < 0 Then
                                Adjacent = y    ' first one found
                            Else
                                If rgnRect(Links(.Ref).k_Neighbor(y).Ref).Left < rgnRect(Links(.Ref).k_Neighbor(Adjacent).Ref).Left Then Adjacent = y
                            End If
                        End If
                    End If
                End If
            Next
        End With
        
        If Adjacent > -1 Then ' we can build a bridge
            ' create the bridge/link
            With Edges(Links(rIndex).k_Neighbor(x).Edge)
                ReDim Preserve .Edge(0 To .Count)
                .Edge(.Count) = Links(Links(rIndex).k_Neighbor(x).Ref).k_Neighbor(Adjacent).Edge
                .Count = .Count + 1
            End With
            With Edges(Links(Links(rIndex).k_Neighbor(x).Ref).k_Neighbor(Adjacent).Edge)
                ReDim Preserve .Edge(0 To .Count)
                .Edge(.Count) = Links(rIndex).k_Neighbor(x).Edge
                .Count = .Count + 1
            End With
            ' done. since we're going left to right, there may be other bridge
            ' possibilities with other rectangles further to the right. But
            ' those will be built when this bridge is next examined. The path
            ' routines allow for a bridge > bridge > bridge path
            Exit For
        End If
    End If
Next

End Sub

Private Sub CreateBridgesBottom(rIndex As Long)

' Exercise in walking the Links Tree
' Build bridges between rectangles horizontally across the bottom of an obstacle
' that separates 2 rectangles. Create a bridge to the nearest right rectangle
' only if both rectangles share a common regional rectangle below them.

' Similar to Vertical links, this routine only moves left to right.

Dim x As Long, y As Long, V As Long, Adjacent As Long

For x = 0 To Links(rIndex).k_Links - 1
    ' check all parent links that are physically higher in the region than this one
    If rgnRect(Links(rIndex).k_Neighbor(x).Ref).Bottom = rgnRect(rIndex).Top Then
        Adjacent = -1
        ' this is a potential bridge. See if any of its links are parallel to target
        For y = 0 To Links(rIndex).k_Links - 1
            ' don't let our target be included here
            If x <> y Then
                If rgnRect(Links(rIndex).k_Neighbor(y).Ref).Bottom = rgnRect(rIndex).Top Then
                    ' only look right >>>>
                    If rgnRect(Links(rIndex).k_Neighbor(y).Ref).Left > rgnRect(Links(rIndex).k_Neighbor(x).Ref).Right Then
                        ' ok, now see if it is physically closer than any previous horizonal neighbor
                        If Adjacent < 0 Then
                            Adjacent = y
                        Else
                            If rgnRect(Links(rIndex).k_Neighbor(y).Ref).Left < rgnRect(Links(rIndex).k_Neighbor(Adjacent).Ref).Left Then Adjacent = y
                        End If
                    End If
                End If
            End If
        Next
        
        If Adjacent > -1 Then ' we have a bridge
            With Edges(Links(rIndex).k_Neighbor(x).Edge)
                ReDim Preserve .Edge(0 To .Count)
                .Edge(.Count) = Links(rIndex).k_Neighbor(Adjacent).Edge
                .Count = .Count + 1
            End With
            With Edges(Links(rIndex).k_Neighbor(Adjacent).Edge)
                ReDim Preserve .Edge(0 To .Count)
                .Edge(.Count) = Links(rIndex).k_Neighbor(x).Edge
                .Count = .Count + 1
            End With
        End If
        ' unlike CreateBridgesTop, this routine will not exit loop at this point
    End If
Next

End Sub

Private Function GetRectNeighbor(rIndex As Long, nIndex As Long, X1 As Long, X2 As Long) As Long
' Simple function that determines if a rectangle in the rgnRect array
' is vertically adjacent to the passed rectangle (rIndex)

Dim lRtn As Long
Dim rID As Long

lRtn = -1
For rID = nIndex To UBound(rgnRect)
    If rgnRect(rID).Top = rgnRect(rIndex).Bottom Then
        If HasCommonEdge(rIndex, rID, X1, X2) Then
            lRtn = rID
            Exit For
        End If
    Else
        If rgnRect(rID).Top > rgnRect(rIndex).Bottom Then Exit For
    End If
Next
GetRectNeighbor = lRtn
End Function

Private Function HasCommonEdge(sourceID As Long, targetID As Long, X1 As Long, X2 As Long) As Boolean
' function returns true if 2 rectangles have a common edge (touch each other)
' vertically. If so, then the far left common point (X1) and far right common
' point (X2) are also returned

Dim tRect As RECT
tRect = rgnRect(targetID)
' must inflate one. If one Rect's Top=other's Bottom, then
'    then IntersectRect returns false
InflateRect tRect, 0, 1
If IntersectRect(tRect, tRect, rgnRect(sourceID)) Then
    X1 = tRect.Left
    X2 = tRect.Right
    HasCommonEdge = True
End If
End Function

Private Function LocateTargets(startPtX As Long, startPtY As Long, _
                                endPtX As Long, endPtY As Long) As Long

' function converts path start & end coordinates into rectangle & edge references

Dim x As Long, tID As Long
Dim y As Long, lat As Long, lng As Long
Dim targetRect As Long, bTargetFound As Byte

For tID = 1 To 2
    ' set the correct coordinates
    If tID = 2 Then
        lat = endPtX
        lng = endPtY
    Else
        lat = startPtX
        lng = startPtY
    End If
    ' determine which rectangle our point falls in
    For x = 0 To UBound(rgnRect)
        If PtInRect(rgnRect(x), lat, lng) Then
            ' update return value
            bTargetFound = bTargetFound + 1
            
            ' build the edge
            With Edges(tID)
                .x = lat
                .Z = lat
                .y = lng
                .Count = 0
            End With
            If UBound(Nodes) Then
                Nodes(tID).Parent = tID
                
                ' create edge links
                For y = 0 To Links(x).k_Links - 1
                    
                    If tID = 2 Then
                    
                        If x = targetRect And bTargetFound = 2 Then
                            ' same regional rectangle, straight line path
                            targetRect = -1
                        Else
                            ' terminating point; becomes sink hole
                            With Edges(Links(x).k_Neighbor(y).Edge)
                                ReDim Preserve .Edge(0 To .Count)
                                .Edge(.Count) = tID
                                .Count = .Count + 1
                            End With
                        End If
                        
                    Else    ' starting point, becomes source only
                        
                        targetRect = x
                        With Edges(tID)
                            If rgnRect(Links(x).k_Neighbor(y).Ref).Top <> rgnRect(x).Top Then
                                If rgnRect(Links(x).k_Neighbor(y).Ref).Bottom <> rgnRect(x).Bottom Then
                                    ReDim Preserve .Edge(0 To .Count)
                                    .Edge(.Count) = Links(x).k_Neighbor(y).Edge
                                    .Count = .Count + 1
                                End If
                            End If
                        End With
                    End If
                Next
            End If
            
            Exit For
            
        End If
    Next
Next

If targetRect < 0 Then
    LocateTargets = 2
Else
    LocateTargets = Abs(bTargetFound = 2)
End If
End Function

