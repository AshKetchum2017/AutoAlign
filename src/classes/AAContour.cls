Option Explicit

' Kurva asli tetap dipakai untuk overlap. Polyline adaptif hanya untuk jarak gap.
' Error tiap chord disimpan dan ditambahkan pada jarak aman, bukan diabaikan.
Private Const AA_FLATNESS_MM As Double = 0.002
Private Const AA_DISTANCE_TOLERANCE_MM As Double = 0.0001
Private Const AA_MAX_SUBDIVISION As Long = 18
Private Const AA_MAX_EDGES As Long = 20000
Private Const AA_INDEX_CELL_MM As Double = 1#
Private Const AA_MAX_CELLS_PER_EDGE As Long = 256
Private Const AA_MAX_QUERY_CELLS As Long = 4096
Private Const AA_INDEX_MIN_EDGES As Long = 32

Private mEdges As Collection
' Data Double dan batas segmen disiapkan sekali; Collection lama untuk metadata.
' Kolom: x1, y1, x2, y2, error, minX, minY, maxX, maxY.
Private mEdgeData() As Double, mEdgeCapacity As Long
Private mEdgeIndexes As Collection
Private mAnchors As Collection
Private mLeft As Double, mTop As Double
Private mMaxError As Double
Private mEdgeBins As Object
Private mWideEdges As Collection
Private mSeen() As Long, mQueryStamp As Long
Private mLastQueryKey As String
Private mLastNearby As Collection
Private mEmptyEdges As Collection
Private mMinX As Double, mMinY As Double, mMaxX As Double, mMaxY As Double

Private Sub Class_Initialize()
    Set mEdges = New Collection
    Set mEdgeIndexes = New Collection
    Set mAnchors = New Collection
    Set mEmptyEdges = New Collection
End Sub

Public Sub Capture(ByVal geometry As AAGeometry)
    Dim part As Curve, edge As Segment, vertex As Node
    Dim i As Long
    Dim x0 As Double, y0 As Double, x1 As Double, y1 As Double
    Dim x2 As Double, y2 As Double, x3 As Double, y3 As Double
    Dim ax As Double, ay As Double, bx As Double, by As Double
    Dim perfStarted As Double
    If AADebugProfiling Then perfStarted = AADebugPerfBegin()
    mLeft = geometry.LeftX: mTop = geometry.TopY
    For i = 1 To geometry.CurveCount
        Set part = geometry.CurveAt(i)
        For Each vertex In part.Nodes
            mAnchors.Add Array(vertex.PositionX - mLeft, mTop - vertex.PositionY)
        Next vertex
        For Each edge In part.Segments
            edge.GetPointPositionAt x0, y0, 0#, cdrParamSegmentOffset
            edge.GetPointPositionAt x3, y3, 1#, cdrParamSegmentOffset
            If edge.Type = cdrLineSegment Then
                AddEdge x0, y0, x3, y3, 0#
            Else
                ' Rekonstruksi kontrol cubic dari empat titik parametrik.
                edge.GetPointPositionAt x1, y1, 1# / 3#, cdrParamSegmentOffset
                edge.GetPointPositionAt x2, y2, 2# / 3#, cdrParamSegmentOffset
                ax = 27# * x1 - 8# * x0 - x3
                ay = 27# * y1 - 8# * y0 - y3
                bx = 27# * x2 - x0 - 8# * x3
                by = 27# * y2 - y0 - 8# * y3
                Flatten x0, y0, (2# * ax - bx) / 18#, (2# * ay - by) / 18#, _
                    (2# * bx - ax) / 18#, (2# * by - ay) / 18#, x3, y3, 0
            End If
        Next edge
    Next i
    If AADebugProfiling Then
        AADebugPerfEnd AA_PERF_CONTOUR, perfStarted
        AADebugContourCaptured mAnchors.Count, mEdges.Count, AA_FLATNESS_MM
    End If
End Sub

Private Sub Flatten(ByVal x0 As Double, ByVal y0 As Double, _
                    ByVal x1 As Double, ByVal y1 As Double, _
                    ByVal x2 As Double, ByVal y2 As Double, _
                    ByVal x3 As Double, ByVal y3 As Double, ByVal depth As Long)
    Dim errorMM As Double, secondError As Double
    Dim px As Double, py As Double
    Dim ax As Double, ay As Double, bx As Double, by As Double
    Dim cx As Double, cy As Double, dx As Double, dy As Double
    Dim ex As Double, ey As Double, mx As Double, my As Double
    errorMM = PointDistance(x1, y1, x0, y0, x3, y3, px, py)
    secondError = PointDistance(x2, y2, x0, y0, x3, y3, px, py)
    If secondError > errorMM Then errorMM = secondError
    If errorMM <= AA_FLATNESS_MM Or depth >= AA_MAX_SUBDIVISION Then
        AddEdge x0, y0, x3, y3, errorMM
        Exit Sub
    End If
    ax = (x0 + x1) / 2#: ay = (y0 + y1) / 2#
    bx = (x1 + x2) / 2#: by = (y1 + y2) / 2#
    cx = (x2 + x3) / 2#: cy = (y2 + y3) / 2#
    dx = (ax + bx) / 2#: dy = (ay + by) / 2#
    ex = (bx + cx) / 2#: ey = (by + cy) / 2#
    mx = (dx + ex) / 2#: my = (dy + ey) / 2#
    Flatten x0, y0, ax, ay, dx, dy, mx, my, depth + 1
    Flatten mx, my, ex, ey, cx, cy, x3, y3, depth + 1
End Sub

Private Sub AddEdge(ByVal x1 As Double, ByVal y1 As Double, _
                    ByVal x2 As Double, ByVal y2 As Double, ByVal errorMM As Double)
    Dim index As Long
    If mEdges.Count >= AA_MAX_EDGES Then Err.Raise vbObjectError + 5240, _
        "AAContour", "Kurva terlalu kompleks untuk pemeriksaan gap Medium."
    index = mEdges.Count + 1
    If index > mEdgeCapacity Then
        If mEdgeCapacity = 0 Then
            mEdgeCapacity = AA_INDEX_MIN_EDGES
            ReDim mEdgeData(0 To 8, 1 To mEdgeCapacity)
        Else
            mEdgeCapacity = mEdgeCapacity * 2
            If mEdgeCapacity > AA_MAX_EDGES Then mEdgeCapacity = AA_MAX_EDGES
            ReDim Preserve mEdgeData(0 To 8, 1 To mEdgeCapacity)
        End If
    End If
    mEdgeData(0, index) = x1: mEdgeData(1, index) = y1
    mEdgeData(2, index) = x2: mEdgeData(3, index) = y2
    mEdgeData(4, index) = errorMM
    mEdgeData(5, index) = MinValue(x1, x2): mEdgeData(7, index) = MaxValue(x1, x2)
    mEdgeData(6, index) = MinValue(y1, y2): mEdgeData(8, index) = MaxValue(y1, y2)
    mEdges.Add Array(x1, y1, x2, y2, errorMM)
    mEdgeIndexes.Add index
    If mEdges.Count = 1 Then
        mMinX = MinValue(x1, x2): mMaxX = MaxValue(x1, x2)
        mMinY = MinValue(y1, y2): mMaxY = MaxValue(y1, y2)
    Else
        mMinX = MinValue(mMinX, MinValue(x1, x2))
        mMaxX = MaxValue(mMaxX, MaxValue(x1, x2))
        mMinY = MinValue(mMinY, MinValue(y1, y2))
        mMaxY = MaxValue(mMaxY, MaxValue(y1, y2))
    End If
    If errorMM > mMaxError Then mMaxError = errorMM
    Set mEdgeBins = Nothing
    Set mLastNearby = Nothing
End Sub

' Indeks hanya menyaring pasangan segmen. Rumus jarak dan error chord tetap
' sama; segmen panjang dan query lebar memakai jalur konservatif tanpa filter.
Private Sub BuildEdgeIndex()
    Dim bucket As Collection, key As String
    Dim i As Long, x As Long, y As Long
    Dim leftCell As Long, rightCell As Long, bottomCell As Long, topCell As Long
    Set mEdgeBins = CreateObject("Scripting.Dictionary")
    Set mWideEdges = New Collection
    ReDim mSeen(1 To mEdges.Count)
    mQueryStamp = 0
    For i = 1 To mEdges.Count
        leftCell = Int(mEdgeData(5, i) / AA_INDEX_CELL_MM)
        rightCell = Int(mEdgeData(7, i) / AA_INDEX_CELL_MM)
        bottomCell = Int(mEdgeData(6, i) / AA_INDEX_CELL_MM)
        topCell = Int(mEdgeData(8, i) / AA_INDEX_CELL_MM)
        If (CDbl(rightCell) - leftCell + 1#) * (CDbl(topCell) - bottomCell + 1#) > _
            AA_MAX_CELLS_PER_EDGE Then
            mWideEdges.Add i
        Else
            For x = leftCell To rightCell
                For y = bottomCell To topCell
                    key = CStr(x) & "|" & CStr(y)
                    If mEdgeBins.Exists(key) Then
                        Set bucket = mEdgeBins(key)
                    Else
                        Set bucket = New Collection
                        mEdgeBins.Add key, bucket
                    End If
                    bucket.Add i
                Next y
            Next x
        End If
    Next i
End Sub

Friend Property Get MaxChordError() As Double
    MaxChordError = mMaxError
End Property

Friend Function IsSeparatedFrom(ByVal leftX As Double, ByVal bottomY As Double, _
                                ByVal rightX As Double, ByVal topY As Double, _
                                ByVal margin As Double) As Boolean
    Dim dx As Double, dy As Double
    dx = MaxValue(0#, MaxValue(mMinX - rightX, leftX - mMaxX))
    dy = MaxValue(0#, MaxValue(mMinY - topY, bottomY - mMaxY))
    IsSeparatedFrom = (dx * dx + dy * dy >= margin * margin)
End Function

Private Function CandidateEdgeIndexes(ByVal leftX As Double, ByVal bottomY As Double, _
                               ByVal rightX As Double, ByVal topY As Double) As Collection
    Dim leftCell As Long, rightCell As Long, bottomCell As Long, topCell As Long
    Dim x As Long, y As Long, index As Variant, key As String, queryKey As String
    Dim result As Collection, bucket As Collection
    If leftX > mMaxX Or rightX < mMinX Or bottomY > mMaxY Or topY < mMinY Then
        Set CandidateEdgeIndexes = mEmptyEdges
        Exit Function
    End If
    If mEdges.Count <= AA_INDEX_MIN_EDGES Then
        Set CandidateEdgeIndexes = mEdgeIndexes
        Exit Function
    End If
    leftCell = Int(leftX / AA_INDEX_CELL_MM): rightCell = Int(rightX / AA_INDEX_CELL_MM)
    bottomCell = Int(bottomY / AA_INDEX_CELL_MM): topCell = Int(topY / AA_INDEX_CELL_MM)
    If (CDbl(rightCell) - leftCell + 1#) * (CDbl(topCell) - bottomCell + 1#) > _
        AA_MAX_QUERY_CELLS Then
        Set CandidateEdgeIndexes = mEdgeIndexes
        Exit Function
    End If
    If mEdgeBins Is Nothing Then BuildEdgeIndex
    queryKey = CStr(leftCell) & "|" & CStr(rightCell) & "|" & _
        CStr(bottomCell) & "|" & CStr(topCell)
    If Not mLastNearby Is Nothing Then
        If queryKey = mLastQueryKey Then
            Set CandidateEdgeIndexes = mLastNearby
            Exit Function
        End If
    End If
    If mQueryStamp >= 2147483646 Then
        ReDim mSeen(1 To mEdges.Count)
        mQueryStamp = 0
    End If
    mQueryStamp = mQueryStamp + 1
    Set result = New Collection
    For Each index In mWideEdges
        result.Add CLng(index)
        mSeen(CLng(index)) = mQueryStamp
    Next index
    For x = leftCell To rightCell
        For y = bottomCell To topCell
            key = CStr(x) & "|" & CStr(y)
            If mEdgeBins.Exists(key) Then
                Set bucket = mEdgeBins(key)
                For Each index In bucket
                    If mSeen(CLng(index)) <> mQueryStamp Then
                        result.Add CLng(index)
                        mSeen(CLng(index)) = mQueryStamp
                    End If
                Next index
            End If
        Next y
    Next x
    mLastQueryKey = queryKey
    Set mLastNearby = result
    Set CandidateEdgeIndexes = result
End Function

Friend Property Get Edges() As Collection
    Set Edges = mEdges
End Property

Public Property Get Anchors() As Collection
    Set Anchors = mAnchors
End Property

Public Property Get LeftX() As Double
    LeftX = mLeft
End Property

Public Property Get TopY() As Double
    TopY = mTop
End Property

Public Function ClearOf(ByVal other As AAContour, ByVal dx As Double, _
                        ByVal dy As Double, ByVal otherDX As Double, _
                        ByVal otherDY As Double, ByVal gapH As Double, _
                        ByVal gapV As Double) As Boolean
    Dim index As Long
    Dim ax As Double, ay As Double, bx As Double, by As Double
    Dim largestGap As Double
    Dim perfStarted As Double
    If AADebugProfiling Then perfStarted = AADebugPerfBegin()
    largestGap = gapH
    If gapV > largestGap Then largestGap = gapV
    If largestGap <= 0# Then
        ClearOf = True
        If AADebugProfiling Then AADebugPerfEnd AA_PERF_GAP, perfStarted
        Exit Function
    End If
    If mEdges.Count = 0 Or other.Edges.Count = 0 Then
        ClearOf = True
        If AADebugProfiling Then AADebugPerfEnd AA_PERF_GAP, perfStarted
        Exit Function
    End If
    If other.IsSeparatedFrom(mMinX + dx - otherDX, mMinY + dy - otherDY, _
        mMaxX + dx - otherDX, mMaxY + dy - otherDY, _
        largestGap + mMaxError + other.MaxChordError) Then
        ClearOf = True
        If AADebugProfiling Then AADebugPerfEnd AA_PERF_GAP, perfStarted
        Exit Function
    End If
    For index = 1 To mEdges.Count
        ax = mEdgeData(0, index) + dx: ay = mEdgeData(1, index) + dy
        bx = mEdgeData(2, index) + dx: by = mEdgeData(3, index) + dy
        If Not other.IsClearOfEdge(ax, ay, bx, by, mEdgeData(4, index), _
            otherDX, otherDY, gapH, gapV, largestGap) Then
            If AADebugProfiling Then AADebugPerfEnd AA_PERF_GAP, perfStarted
            Exit Function
        End If
    Next index
    ClearOf = True
    If AADebugProfiling Then AADebugPerfEnd AA_PERF_GAP, perfStarted
End Function

' Jalankan loop pada data Double milik kontur tujuan agar koordinat/batas
' tidak disalin sebagai Variant untuk setiap pasangan segmen.
Friend Function IsClearOfEdge(ByVal ax As Double, ByVal ay As Double, _
                             ByVal bx As Double, ByVal by As Double, _
                             ByVal errorMM As Double, ByVal dx As Double, ByVal dy As Double, _
                             ByVal gapH As Double, ByVal gapV As Double, _
                             ByVal largestGap As Double) As Boolean
    Dim nearby As Collection, item As Variant, index As Long
    Dim minX As Double, minY As Double, maxX As Double, maxY As Double
    Dim cx As Double, cy As Double, ex As Double, ey As Double
    Dim searchAllowance As Double, allowance As Double, edgeError As Double
    Dim vx As Double, vy As Double, distance As Double, gap As Double
    minX = MinValue(ax, bx): maxX = MaxValue(ax, bx)
    minY = MinValue(ay, by): maxY = MaxValue(ay, by)
    searchAllowance = largestGap + errorMM + mMaxError
    Set nearby = CandidateEdgeIndexes(minX - dx - searchAllowance, _
        minY - dy - searchAllowance, maxX - dx + searchAllowance, _
        maxY - dy + searchAllowance)
    For Each item In nearby
        index = CLng(item)
        edgeError = mEdgeData(4, index)
        allowance = largestGap + errorMM + edgeError
        If minX <= mEdgeData(7, index) + dx + allowance And _
           maxX >= mEdgeData(5, index) + dx - allowance And _
           minY <= mEdgeData(8, index) + dy + allowance And _
           maxY >= mEdgeData(6, index) + dy - allowance Then
            cx = mEdgeData(0, index) + dx: cy = mEdgeData(1, index) + dy
            ex = mEdgeData(2, index) + dx: ey = mEdgeData(3, index) + dy
            distance = EdgeDistance(ax, ay, bx, by, cx, cy, ex, ey, vx, vy)
            gap = largestGap
            ' Normal horizontal -> H; normal vertikal -> V; miring -> max(H,V).
            If Abs(vy) <= AA_DISTANCE_TOLERANCE_MM Then
                gap = gapH
            ElseIf Abs(vx) <= AA_DISTANCE_TOLERANCE_MM Then
                gap = gapV
            End If
            If gap > 0# Then
                If distance < gap + errorMM + edgeError - AA_DISTANCE_TOLERANCE_MM Then
                    Exit Function
                End If
            End If
        End If
    Next item
    IsClearOfEdge = True
End Function

Private Function EdgeDistance(ByVal ax As Double, ByVal ay As Double, _
                              ByVal bx As Double, ByVal by As Double, _
                              ByVal cx As Double, ByVal cy As Double, _
                              ByVal dx As Double, ByVal dy As Double, _
                              ByRef vx As Double, ByRef vy As Double) As Double
    Dim denominator As Double, t As Double, u As Double
    Dim best As Double, value As Double, px As Double, py As Double
    denominator = (bx - ax) * (dy - cy) - (by - ay) * (dx - cx)
    If Abs(denominator) > 0.000000000001 Then
        t = ((cx - ax) * (dy - cy) - (cy - ay) * (dx - cx)) / denominator
        u = ((cx - ax) * (by - ay) - (cy - ay) * (bx - ax)) / denominator
        If t >= 0# And t <= 1# And u >= 0# And u <= 1# Then
            vx = 0#: vy = 0#
            Exit Function
        End If
    End If
    best = PointDistance(ax, ay, cx, cy, dx, dy, px, py)
    vx = ax - px: vy = ay - py
    value = PointDistance(bx, by, cx, cy, dx, dy, px, py)
    If value < best Then best = value: vx = bx - px: vy = by - py
    value = PointDistance(cx, cy, ax, ay, bx, by, px, py)
    If value < best Then best = value: vx = px - cx: vy = py - cy
    value = PointDistance(dx, dy, ax, ay, bx, by, px, py)
    If value < best Then best = value: vx = px - dx: vy = py - dy
    EdgeDistance = best
End Function

Private Function PointDistance(ByVal x As Double, ByVal y As Double, _
                               ByVal ax As Double, ByVal ay As Double, _
                               ByVal bx As Double, ByVal by As Double, _
                               ByRef px As Double, ByRef py As Double) As Double
    Dim length2 As Double, t As Double
    length2 = (bx - ax) ^ 2 + (by - ay) ^ 2
    If length2 > 0# Then t = ((x - ax) * (bx - ax) + (y - ay) * (by - ay)) / length2
    If t < 0# Then t = 0#
    If t > 1# Then t = 1#
    px = ax + t * (bx - ax): py = ay + t * (by - ay)
    PointDistance = Sqr((x - px) ^ 2 + (y - py) ^ 2)
End Function

Private Function MinValue(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinValue = a Else MinValue = b
End Function

Private Function MaxValue(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxValue = a Else MaxValue = b
End Function
