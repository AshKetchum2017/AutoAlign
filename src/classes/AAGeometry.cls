Option Explicit

' Salinan kurva bebas untuk perencanaan; tidak mengubah Shape sumber.
Private Const AA_GEOMETRY_TOLERANCE_MM As Double = 0.0001
Private Const AA_AREA_TOLERANCE_MM2 As Double = 0.000001
Private Const AA_RECTANGLE_DIRECTION_TOLERANCE As Double = 0.000001

Private mCurves As Collection
Private mLeft As Double
Private mBottom As Double
Private mRight As Double
Private mTop As Double
Private mHasBounds As Boolean

Private Sub Class_Initialize()
    Set mCurves = New Collection
End Sub

Friend Property Get CurveCount() As Long
    CurveCount = mCurves.Count
End Property

Friend Property Get CurveAt(ByVal index As Long) As Curve
    Set CurveAt = mCurves(index)
End Property

Public Sub Capture(ByVal sourceShape As Shape)
    AddShapeCurves sourceShape
    If mCurves.Count = 0 Then Err.Raise vbObjectError + 5220, _
        "AAGeometry", "Object tidak memiliki kurva untuk pemeriksaan collision."
End Sub

' Periksa path, sehingga rectangle yang sudah menjadi Curve juga dikenali.
Public Function IsSharpRectangle() As Boolean
    Dim part As Curve
    Dim path As SubPath
    Dim edge As Segment
    Dim dx() As Double, dy() As Double
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim vx As Double, vy As Double, length As Double
    Dim sideCount As Long
    Dim i As Long, nextSide As Long

    If mCurves.Count <> 1 Then Exit Function
    Set part = mCurves(1)
    If part.SubPaths.Count <> 1 Then Exit Function
    Set path = part.SubPaths(1)
    If Not path.Closed Or path.Segments.Count < 4 Then Exit Function
    ReDim dx(1 To path.Segments.Count)
    ReDim dy(1 To path.Segments.Count)
    For Each edge In path.Segments
        If edge.Type <> cdrLineSegment Then Exit Function
        edge.GetPointPositionAt x1, y1, 0#, cdrParamSegmentOffset
        edge.GetPointPositionAt x2, y2, 1#, cdrParamSegmentOffset
        vx = x2 - x1
        vy = y2 - y1
        length = Sqr(vx * vx + vy * vy)
        If length <= AA_GEOMETRY_TOLERANCE_MM Then Exit Function
        vx = vx / length
        vy = vy / length
        If sideCount = 0 Then
            sideCount = 1
        ElseIf Not SameDirection(dx(sideCount), dy(sideCount), vx, vy) Then
            sideCount = sideCount + 1
        End If
        dx(sideCount) = vx
        dy(sideCount) = vy
    Next edge
    ' Node awal boleh berada di tengah salah satu sisi.
    If sideCount > 1 Then
        If SameDirection(dx(1), dy(1), dx(sideCount), dy(sideCount)) Then _
            sideCount = sideCount - 1
    End If
    If sideCount <> 4 Then Exit Function
    For i = 1 To 4
        nextSide = i Mod 4 + 1
        If Abs(dx(i) * dx(nextSide) + dy(i) * dy(nextSide)) > _
            AA_RECTANGLE_DIRECTION_TOLERANCE Then Exit Function
        ' Empat belokan harus searah; menolak path yang berbalik/bersilangan.
        If i > 1 Then
            If (dx(i) * dy(nextSide) - dy(i) * dx(nextSide)) * _
                (dx(1) * dy(2) - dy(1) * dx(2)) <= 0# Then Exit Function
        End If
    Next i
    IsSharpRectangle = True
End Function

Public Function IsSharpTriangle() As Boolean
    Dim part As Curve, path As SubPath, edge As Segment
    Dim dx() As Double, dy() As Double
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim vx As Double, vy As Double, length As Double, turn As Double
    Dim sideCount As Long, i As Long, nextSide As Long

    If mCurves.Count <> 1 Then Exit Function
    Set part = mCurves(1)
    If part.SubPaths.Count <> 1 Then Exit Function
    Set path = part.SubPaths(1)
    If Not path.Closed Or path.Segments.Count < 3 Then Exit Function
    ReDim dx(1 To path.Segments.Count)
    ReDim dy(1 To path.Segments.Count)
    For Each edge In path.Segments
        If edge.Type <> cdrLineSegment Then Exit Function
        edge.GetPointPositionAt x1, y1, 0#, cdrParamSegmentOffset
        edge.GetPointPositionAt x2, y2, 1#, cdrParamSegmentOffset
        vx = x2 - x1: vy = y2 - y1
        length = Sqr(vx * vx + vy * vy)
        If length <= AA_GEOMETRY_TOLERANCE_MM Then Exit Function
        vx = vx / length: vy = vy / length
        If sideCount = 0 Then
            sideCount = 1
        ElseIf Not SameDirection(dx(sideCount), dy(sideCount), vx, vy) Then
            sideCount = sideCount + 1
        End If
        dx(sideCount) = vx: dy(sideCount) = vy
    Next edge
    If sideCount > 1 Then
        If SameDirection(dx(1), dy(1), dx(sideCount), dy(sideCount)) Then _
            sideCount = sideCount - 1
    End If
    If sideCount <> 3 Then Exit Function
    turn = dx(1) * dy(2) - dy(1) * dx(2)
    If Abs(turn) <= AA_RECTANGLE_DIRECTION_TOLERANCE Then Exit Function
    For i = 2 To 3
        nextSide = i Mod 3 + 1
        If (dx(i) * dy(nextSide) - dy(i) * dx(nextSide)) * turn <= 0# Then Exit Function
    Next i
    IsSharpTriangle = True
End Function

Private Function SameDirection(ByVal x1 As Double, ByVal y1 As Double, _
                               ByVal x2 As Double, ByVal y2 As Double) As Boolean
    SameDirection = (Abs(x1 * y2 - y1 * x2) <= _
        AA_RECTANGLE_DIRECTION_TOLERANCE And x1 * x2 + y1 * y2 > 0#)
End Function

Public Function IsRoundedRectangle() As Boolean
    Dim part As Curve
    Dim path As SubPath
    Dim edge As Segment
    Dim dx(1 To 4) As Double, dy(1 To 4) As Double
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim vx As Double, vy As Double, length As Double
    Dim sides As Long, bends As Long, i As Long, nextSide As Long
    Dim turn As Double

    If mCurves.Count <> 1 Then Exit Function
    Set part = mCurves(1)
    If part.SubPaths.Count <> 1 Then Exit Function
    Set path = part.SubPaths(1)
    If Not path.Closed Then Exit Function
    For Each edge In path.Segments
        If edge.Type = cdrLineSegment Then
            edge.GetPointPositionAt x1, y1, 0#, cdrParamSegmentOffset
            edge.GetPointPositionAt x2, y2, 1#, cdrParamSegmentOffset
            vx = x2 - x1
            vy = y2 - y1
            length = Sqr(vx * vx + vy * vy)
            If length <= AA_GEOMETRY_TOLERANCE_MM Then Exit Function
            vx = vx / length
            vy = vy / length
            If sides > 0 Then
                If SameDirection(dx(sides), dy(sides), vx, vy) Then GoTo NextEdge
                If sides = 4 Then
                    If SameDirection(dx(1), dy(1), vx, vy) Then GoTo NextEdge
                    Exit Function
                End If
            End If
            sides = sides + 1
            dx(sides) = vx
            dy(sides) = vy
        Else
            bends = bends + 1
        End If
NextEdge:
    Next edge
    If sides <> 4 Or bends = 0 Then Exit Function
    turn = dx(1) * dy(2) - dy(1) * dx(2)
    For i = 1 To 4
        nextSide = i Mod 4 + 1
        If Abs(dx(i) * dx(nextSide) + dy(i) * dy(nextSide)) > _
            AA_RECTANGLE_DIRECTION_TOLERANCE Then Exit Function
        If (dx(i) * dy(nextSide) - dy(i) * dx(nextSide)) * turn <= 0# Then Exit Function
    Next i
    ' Kurva sudut harus membelok searah sisi, tanpa cekungan/gelombang.
    For Each edge In path.Segments
        If edge.Type <> cdrLineSegment Then
            If Not ConvexCorner(edge, turn) Then Exit Function
        End If
    Next edge
    IsRoundedRectangle = True
End Function

Private Function ConvexCorner(ByVal edge As Segment, ByVal turn As Double) As Boolean
    Dim ax As Double, ay As Double, bx As Double, by As Double
    Dim cx As Double, cy As Double
    Dim i As Long
    edge.GetPointPositionAt ax, ay, 0#, cdrParamSegmentOffset
    edge.GetPointPositionAt bx, by, 0.125, cdrParamSegmentOffset
    For i = 2 To 8
        edge.GetPointPositionAt cx, cy, i / 8#, cdrParamSegmentOffset
        If ((bx - ax) * (cy - by) - (by - ay) * (cx - bx)) * turn < _
            -AA_AREA_TOLERANCE_MM2 Then Exit Function
        ax = bx: ay = by
        bx = cx: by = cy
    Next i
    ConvexCorner = True
End Function

Public Function SameForm(ByVal other As AAGeometry) As Boolean
    Dim first As AAGeometry, second As AAGeometry
    Set first = UnitSizedCopy
    Set second = other.UnitSizedCopy
    SameForm = first.SameGeometry(second)
End Function

' Siluet tertutup mengabaikan artwork yang berada di dalam background Design.
' Weld hanya dilakukan pada salinan kurva bebas, tanpa Ungroup/ubah Shape sumber.
Public Function OuterBoundary() As AAGeometry
    Dim joined As Curve
    Dim part As Curve
    Dim result As AAGeometry

    For Each part In mCurves
        If IsClosedCurve(part) Then
            If joined Is Nothing Then
                Set joined = part.GetCopy
            Else
                Set joined = joined.WeldWith(part)
            End If
        End If
    Next part
    If joined Is Nothing Then
        Set OuterBoundary = Me
        Exit Function
    End If
    Set result = New AAGeometry
    result.AddCurve joined
    ' Kurva terbuka yang menentukan batas luar tetap memerlukan geometri lengkap.
    If Abs(result.LeftX - LeftX) > AA_GEOMETRY_TOLERANCE_MM Or _
       Abs(result.RightX - RightX) > AA_GEOMETRY_TOLERANCE_MM Or _
       Abs(result.BottomY - BottomY) > AA_GEOMETRY_TOLERANCE_MM Or _
       Abs(result.TopY - TopY) > AA_GEOMETRY_TOLERANCE_MM Then
        Set OuterBoundary = Me
    Else
        Set OuterBoundary = result
    End If
End Function

Friend Function UnitSizedCopy() As AAGeometry
    Set UnitSizedCopy = SizedCopy(1#, 1#)
End Function

Public Function SizedCopy(ByVal targetWidth As Double, ByVal targetHeight As Double) As AAGeometry
    Dim result As AAGeometry
    Dim part As Curve, copy As Curve
    Set result = New AAGeometry
    If Width <= 0# Or Height <= 0# Then Err.Raise vbObjectError + 5221, _
        "AAGeometry", "Ukuran kurva tidak dapat dinormalisasi."
    For Each part In mCurves
        Set copy = part.GetCopy
        copy.Nodes.All.Stretch CSng(targetWidth / Width), CSng(targetHeight / Height), _
            True, mLeft, mBottom
        result.AddCurve copy
    Next part
    Set SizedCopy = result.Placed(0#, result.Height)
End Function

' Bounding box pasangan tanpa membaca kurva artwork Design.
Public Sub CaptureBounds(ByVal sourceShape As Shape)
    Dim x As Double, y As Double, widthMM As Double, heightMM As Double
    sourceShape.GetBoundingBox x, y, widthMM, heightMM, False
    If widthMM <= 0# Or heightMM <= 0# Then Err.Raise vbObjectError + 5220, _
        "AAGeometry", "Object tidak memiliki bounding box yang valid."
    AddBounds x, y, x + widthMM, y + heightMM
End Sub

Public Sub AddBounds(ByVal leftX As Double, ByVal bottomY As Double, _
                     ByVal rightX As Double, ByVal topY As Double)
    If Not mHasBounds Then
        mLeft = leftX
        mBottom = bottomY
        mRight = rightX
        mTop = topY
        mHasBounds = True
    Else
        If leftX < mLeft Then mLeft = leftX
        If bottomY < mBottom Then mBottom = bottomY
        If rightX > mRight Then mRight = rightX
        If topY > mTop Then mTop = topY
    End If
End Sub

Public Sub AddRectangle(ByVal leftX As Double, ByVal bottomY As Double, _
                        ByVal rightX As Double, ByVal topY As Double, _
                        ByVal sourceDocument As Document)
    Dim rectangleCurve As Curve
    Dim path As SubPath

    Set rectangleCurve = Application.CreateCurve(sourceDocument)
    Set path = rectangleCurve.CreateSubPath(leftX, bottomY)
    path.AppendLineSegment rightX, bottomY
    path.AppendLineSegment rightX, topY
    path.AppendLineSegment leftX, topY
    path.Closed = True
    AddCurve rectangleCurve
End Sub

Public Sub AddPlaced(ByVal source As AAGeometry, ByVal leftX As Double, _
                     ByVal topY As Double)
    Dim part As Curve
    Dim moved As Curve
    Dim vertex As Node
    Dim deltaX As Double
    Dim deltaY As Double
    Dim i As Long

    deltaX = leftX - source.LeftX
    deltaY = topY - source.TopY
    If source.CurveCount = 0 Then
        AddBounds leftX, topY - source.Height, leftX + source.Width, topY
        Exit Sub
    End If
    For i = 1 To source.CurveCount
        Set part = source.CurveAt(i)
        Set moved = part.GetCopy
        For Each vertex In moved.Nodes
            vertex.Move deltaX, deltaY
        Next vertex
        AddCurve moved
    Next i
End Sub

Public Function Placed(ByVal leftX As Double, ByVal topY As Double) As AAGeometry
    Dim result As AAGeometry
    Set result = New AAGeometry
    result.AddPlaced Me, leftX, topY
    Set Placed = result
End Function

Public Function Rotated(ByVal angle As Double, ByVal centerX As Double, _
                        ByVal centerY As Double) As AAGeometry
    Dim result As AAGeometry
    Dim part As Curve, copy As Curve
    Dim radians As Double, cosine As Double, sine As Double
    Dim x As Double, y As Double, rx As Double, ry As Double
    Dim i As Long
    Set result = New AAGeometry
    If mCurves.Count > 0 Then
        For Each part In mCurves
            Set copy = part.GetCopy
            If angle <> 0# Then copy.Nodes.All.Rotate angle, True, centerX, centerY
            result.AddCurve copy
        Next part
    Else
        radians = angle * 4# * Atn(1#) / 180#
        cosine = Cos(radians): sine = Sin(radians)
        For i = 0 To 3
            If i Mod 2 = 0 Then x = mLeft Else x = mRight
            If i < 2 Then y = mBottom Else y = mTop
            rx = centerX + (x - centerX) * cosine - (y - centerY) * sine
            ry = centerY + (x - centerX) * sine + (y - centerY) * cosine
            result.AddBounds rx, ry, rx, ry
        Next i
    End If
    Set Rotated = result
End Function

Public Function Overlaps(ByVal other As AAGeometry) As Boolean
    Dim firstCurve As Curve
    Dim secondCurve As Curve
    Dim i As Long

    If Not BoundsMayOverlap(other) Then Exit Function
    If mCurves.Count = 0 Or other.CurveCount = 0 Then
        Overlaps = True
        Exit Function
    End If
    For Each firstCurve In mCurves
        For i = 1 To other.CurveCount
            Set secondCurve = other.CurveAt(i)
            If CurvesOverlap(firstCurve, secondCurve) Then
                Overlaps = True
                Exit Function
            End If
        Next i
    Next firstCurve
End Function

Public Function SameGeometry(ByVal other As AAGeometry) As Boolean
    Dim firstNormalized As AAGeometry
    Dim secondNormalized As AAGeometry
    Dim firstCurve As Curve
    Dim secondCurve As Curve
    Dim matched() As Boolean
    Dim i As Long
    Dim firstIndex As Long
    Dim found As Boolean

    If Abs(Width - other.Width) > AA_GEOMETRY_TOLERANCE_MM Or _
        Abs(Height - other.Height) > AA_GEOMETRY_TOLERANCE_MM Then Exit Function
    If mCurves.Count <> other.CurveCount Then Exit Function
    If mCurves.Count = 0 Then
        SameGeometry = True
        Exit Function
    End If
    Set firstNormalized = Placed(0#, Height)
    Set secondNormalized = other.Placed(0#, other.Height)
    ReDim matched(1 To mCurves.Count)
    For firstIndex = 1 To firstNormalized.CurveCount
        Set firstCurve = firstNormalized.CurveAt(firstIndex)
        found = False
        For i = 1 To secondNormalized.CurveCount
            If Not matched(i) Then
                Set secondCurve = secondNormalized.CurveAt(i)
                If IsClosedCurve(firstCurve) = IsClosedCurve(secondCurve) Then
                    If IsClosedCurve(firstCurve) Then
                        found = CurvesSameArea(firstCurve, secondCurve)
                    Else
                        found = CurvesSameTrace(firstCurve, secondCurve)
                    End If
                    If found Then
                        matched(i) = True
                        Exit For
                    End If
                End If
            End If
        Next i
        If Not found Then Exit Function
    Next firstIndex
    SameGeometry = True
End Function

Private Function CurvesSameTrace(ByVal firstCurve As Curve, _
                                 ByVal secondCurve As Curve) As Boolean
    Dim firstSegment As Segment
    Dim secondSegment As Segment
    Dim i As Long
    Dim sampleIndex As Long
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim sampleOffset As Double

    If firstCurve.Segments.Count <> secondCurve.Segments.Count Then Exit Function
    For i = 1 To firstCurve.Segments.Count
        Set firstSegment = firstCurve.Segments(i)
        Set secondSegment = secondCurve.Segments(i)
        For sampleIndex = 0 To 4
            sampleOffset = sampleIndex / 4#
            firstSegment.GetPointPositionAt x1, y1, sampleOffset, _
                cdrParamSegmentOffset
            secondSegment.GetPointPositionAt x2, y2, sampleOffset, _
                cdrParamSegmentOffset
            If Abs(x1 - x2) > AA_GEOMETRY_TOLERANCE_MM Or _
               Abs(y1 - y2) > AA_GEOMETRY_TOLERANCE_MM Then Exit Function
        Next sampleIndex
    Next i
    CurvesSameTrace = True
End Function

Public Property Get LeftX() As Double
    LeftX = mLeft
End Property

Public Property Get RightX() As Double
    RightX = mRight
End Property

Public Property Get BottomY() As Double
    BottomY = mBottom
End Property

Public Property Get TopY() As Double
    TopY = mTop
End Property

Public Property Get Width() As Double
    Width = mRight - mLeft
End Property

Public Property Get Height() As Double
    Height = mTop - mBottom
End Property

Private Sub AddShapeCurves(ByVal sourceShape As Shape)
    Dim child As Shape
    Dim display As Curve

    If sourceShape.Type = cdrGroupShape Then
        For Each child In sourceShape.Shapes
            AddShapeCurves child
        Next child
    Else
        Set display = sourceShape.DisplayCurve
        If display Is Nothing Then Err.Raise vbObjectError + 5220, _
            "AAGeometry", "Salah satu anggota object tidak memiliki kurva."
        AddCurve display.GetCopy
    End If
End Sub

Friend Sub AddCurve(ByVal part As Curve)
    Dim x As Double
    Dim y As Double
    Dim widthMM As Double
    Dim heightMM As Double

    If Not part.GetBoundingBox(x, y, widthMM, heightMM) Then Err.Raise _
        vbObjectError + 5220, "AAGeometry", _
        "Kurva object tidak memiliki bounding box yang valid."
    mCurves.Add part
    If Not mHasBounds Then
        mLeft = x
        mBottom = y
        mRight = x + widthMM
        mTop = y + heightMM
        mHasBounds = True
    Else
        If x < mLeft Then mLeft = x
        If y < mBottom Then mBottom = y
        If x + widthMM > mRight Then mRight = x + widthMM
        If y + heightMM > mTop Then mTop = y + heightMM
    End If
End Sub

Private Function BoundsMayOverlap(ByVal other As AAGeometry) As Boolean
    BoundsMayOverlap = (mLeft < other.RightX - AA_GEOMETRY_TOLERANCE_MM And _
        mRight > other.LeftX + AA_GEOMETRY_TOLERANCE_MM And _
        mBottom < other.TopY - AA_GEOMETRY_TOLERANCE_MM And _
        mTop > other.BottomY + AA_GEOMETRY_TOLERANCE_MM)
End Function

Private Function CurvesOverlap(ByVal firstCurve As Curve, _
                               ByVal secondCurve As Curve) As Boolean
    Dim joined As Curve
    Dim overlapArea As Double

    If IsClosedCurve(firstCurve) And IsClosedCurve(secondCurve) Then
        Set joined = firstCurve.WeldWith(secondCurve)
        overlapArea = Abs(firstCurve.Area) + Abs(secondCurve.Area) - Abs(joined.Area)
        CurvesOverlap = (overlapArea > AA_AREA_TOLERANCE_MM2)
    ElseIf IsClosedCurve(firstCurve) Then
        CurvesOverlap = OpenCurveEntersClosed(secondCurve, firstCurve)
    ElseIf IsClosedCurve(secondCurve) Then
        CurvesOverlap = OpenCurveEntersClosed(firstCurve, secondCurve)
    Else
        CurvesOverlap = OpenCurvesCross(firstCurve, secondCurve)
    End If
End Function

Private Function CurvesSameArea(ByVal firstCurve As Curve, _
                                ByVal secondCurve As Curve) As Boolean
    Dim joined As Curve
    Dim commonArea As Double
    Set joined = firstCurve.WeldWith(secondCurve)
    commonArea = Abs(firstCurve.Area) + Abs(secondCurve.Area) - Abs(joined.Area)
    CurvesSameArea = (Abs(Abs(firstCurve.Area) - commonArea) <= AA_AREA_TOLERANCE_MM2 _
        And Abs(Abs(secondCurve.Area) - commonArea) <= AA_AREA_TOLERANCE_MM2)
End Function

Private Function IsClosedCurve(ByVal sourceCurve As Curve) As Boolean
    Dim path As SubPath
    If sourceCurve.SubPaths.Count = 0 Then Exit Function
    For Each path In sourceCurve.SubPaths
        If Not path.Closed Then Exit Function
    Next path
    IsClosedCurve = True
End Function

Private Function OpenCurveEntersClosed(ByVal openCurve As Curve, _
                                       ByVal closedCurve As Curve) As Boolean
    Dim vertex As Node
    Dim openSegment As Segment
    Dim closedSegment As Segment
    Dim crossings As CrossPoints
    Dim crossing As CrossPoint

    For Each vertex In openCurve.Nodes
        If closedCurve.IsPointInside(vertex.PositionX, vertex.PositionY) Then
            OpenCurveEntersClosed = True
            Exit Function
        End If
    Next vertex
    For Each openSegment In openCurve.Segments
        For Each closedSegment In closedCurve.Segments
            Set crossings = openSegment.GetIntersections( _
                closedSegment, cdrParamSegmentOffset)
            For Each crossing In crossings
                If PointBesideInside(openSegment, crossing.Offset, closedCurve) Then
                    OpenCurveEntersClosed = True
                    Exit Function
                End If
            Next crossing
        Next closedSegment
    Next openSegment
End Function

Private Function PointBesideInside(ByVal path As Segment, ByVal offset As Double, _
                                   ByVal region As Curve) As Boolean
    Dim x As Double
    Dim y As Double
    Dim testOffset As Double

    testOffset = offset - 0.00001
    If testOffset > 0# Then
        path.GetPointPositionAt x, y, testOffset, cdrParamSegmentOffset
        If region.IsPointInside(x, y) Then
            PointBesideInside = True
            Exit Function
        End If
    End If
    testOffset = offset + 0.00001
    If testOffset < 1# Then
        path.GetPointPositionAt x, y, testOffset, cdrParamSegmentOffset
        PointBesideInside = region.IsPointInside(x, y)
    End If
End Function

Private Function OpenCurvesCross(ByVal firstCurve As Curve, _
                                 ByVal secondCurve As Curve) As Boolean
    Dim firstSegment As Segment
    Dim secondSegment As Segment
    Dim crossings As CrossPoints
    Dim crossing As CrossPoint
    Dim ax As Double, ay As Double, bx As Double, by As Double
    Dim cx As Double, cy As Double, dx As Double, dy As Double
    Dim crossProduct As Double

    For Each firstSegment In firstCurve.Segments
        For Each secondSegment In secondCurve.Segments
            Set crossings = firstSegment.GetIntersections( _
                secondSegment, cdrParamSegmentOffset)
            For Each crossing In crossings
                firstSegment.GetPointPositionAt ax, ay, _
                    ClampOffset(crossing.Offset - 0.00001), cdrParamSegmentOffset
                firstSegment.GetPointPositionAt bx, by, _
                    ClampOffset(crossing.Offset + 0.00001), cdrParamSegmentOffset
                secondSegment.GetPointPositionAt cx, cy, _
                    ClampOffset(crossing.Offset2 - 0.00001), cdrParamSegmentOffset
                secondSegment.GetPointPositionAt dx, dy, _
                    ClampOffset(crossing.Offset2 + 0.00001), cdrParamSegmentOffset
                crossProduct = (bx - ax) * (dy - cy) - (by - ay) * (dx - cx)
                If Abs(crossProduct) > AA_AREA_TOLERANCE_MM2 Then
                    OpenCurvesCross = True
                    Exit Function
                End If
            Next crossing
        Next secondSegment
    Next firstSegment
End Function

Private Function ClampOffset(ByVal value As Double) As Double
    If value < 0# Then
        ClampOffset = 0#
    ElseIf value > 1# Then
        ClampOffset = 1#
    Else
        ClampOffset = value
    End If
End Function
