Option Explicit

' Eksperimental: hanya pencocokan posisi relatif, bukan toleransi clearance.
Private Const AA_OPTIMIZE_MATCH_TOLERANCE_MM As Double = 1#
Private Const AA_OPTIMIZE_ANGLE_TOLERANCE As Double = 0.000001
Private Const AA_EDGE_TOLERANCE_MM As Double = 0.0001

Private mDesigns As Collection, mCuts As Collection
Private mCurrent() As AAPlacement, mOriginal() As AAPlacement
Private mDesignX() As Double, mDesignY() As Double, mDesignAngle() As Double
Private mCutX() As Double, mCutY() As Double, mCutAngle() As Double
Private mMatchX() As Double, mMatchY() As Double
Private mMoving() As Boolean
Private mEnvelopeLeft As Double, mEnvelopeBottom As Double
Private mEnvelopeRight As Double, mEnvelopeTop As Double
Private mCount As Long
Private mDesignOriginX As Double, mDesignOriginY As Double
Private mCutOriginX As Double, mCutOriginY As Double
Private mSettings As AASettings, mForbidden As Collection
Private mOccupied As Collection, mIndex As AAOccupiedIndex
Private mPageCenterX As Double, mPageCenterY As Double

Public Sub Prepare(ByVal designRoot As Shape, ByVal cutRoot As Shape, _
                   ByVal settings As AASettings, ByVal forbidden As Collection, _
                   Optional ByVal pageCenterX As Double = 0#, Optional ByVal pageCenterY As Double = 0#)
    Dim designs As Collection, cuts As Collection, used() As Boolean
    Dim design As Shape, cut As Shape
    Dim dOriginX As Double, dOriginY As Double, cOriginX As Double, cOriginY As Double
    Dim designGeometry As AAGeometry, collision As AAGeometry
    Dim orientation As AARotationVariant, box As AAGeometry
    Dim i As Long, j As Long, found As Long, matches As Long
    Dim h As Double, v As Double

    Set mSettings = settings: Set mForbidden = forbidden
    mPageCenterX = pageCenterX: mPageCenterY = pageCenterY
    Set mIndex = New AAOccupiedIndex
    Set designs = Members(designRoot)
    AreaOrigin designRoot, dOriginX, dOriginY
    mDesignOriginX = dOriginX: mDesignOriginY = dOriginY
    Set mDesigns = New Collection: Set mCuts = New Collection
    mCount = designs.Count
    If mCount = 0 Then Fail "Container tidak memiliki objek."
    ReDim mCurrent(1 To mCount): ReDim mOriginal(1 To mCount)
    ReDim mDesignX(1 To mCount): ReDim mDesignY(1 To mCount)
    ReDim mDesignAngle(1 To mCount)
    ReDim mCutX(1 To mCount): ReDim mCutY(1 To mCount)
    ReDim mCutAngle(1 To mCount)
    ReDim mMatchX(1 To mCount): ReDim mMatchY(1 To mCount)
    ReDim mMoving(1 To mCount)
    If Not cutRoot Is Nothing Then
        Set cuts = Members(cutRoot)
        If cuts.Count <> mCount Then Fail "Jumlah anggota Design dan Cut Line harus sama."
        AreaOrigin cutRoot, cOriginX, cOriginY
        mCutOriginX = cOriginX: mCutOriginY = cOriginY
        ReDim used(1 To mCount)
    End If

    AAProgressStage "Membaca susunan", "Geometri dan pasangan", 0, mCount
    For i = 1 To mCount
        AAProgressStep i - 1, mCount
        Set cut = Nothing
        If cuts Is Nothing Then
            Set design = designs(i)
        Else
            ' Urutan Cut Line menjadi indeks bersama untuk seluruh template.
            Set cut = cuts(i)
            matches = 0: found = 0
            For j = 1 To mCount
                If PositionMatches(designs(j), dOriginX, dOriginY, _
                                   cut, cOriginX, cOriginY) Then
                    matches = matches + 1: found = j
                End If
            Next j
            If matches <> 1 Then Fail "Pasangan Design/Cut Line tidak unik atau di luar toleransi 1 mm."
            If used(found) Then Fail "Satu Design terpetakan ke lebih dari satu Cut Line."
            used(found) = True
            Set design = designs(found)
        End If
        mDesigns.Add design
        mDesignX(i) = design.LeftX: mDesignY(i) = design.TopY
        mDesignAngle(i) = design.RotationAngle
        Set designGeometry = New AAGeometry
        If cut Is Nothing Then
            designGeometry.Capture design
            Set designGeometry = designGeometry.OuterBoundary
            Set designGeometry = designGeometry.Placed(designGeometry.LeftX - dOriginX, _
                designGeometry.TopY - dOriginY)
            Set collision = designGeometry
        Else
            mCuts.Add cut
            mCutX(i) = cut.LeftX: mCutY(i) = cut.TopY
            mCutAngle(i) = cut.RotationAngle
            If settings.IncludeDesignZones Then
                designGeometry.Capture design
                Set designGeometry = designGeometry.OuterBoundary
            Else
                designGeometry.CaptureBounds design
            End If
            Set designGeometry = designGeometry.Placed(designGeometry.LeftX - dOriginX, _
                designGeometry.TopY - dOriginY)
            Set collision = New AAGeometry
            collision.Capture cut
            Set collision = collision.Placed(collision.LeftX - cOriginX, collision.TopY - cOriginY)
        End If
        h = settings.DefaultHorizontalGapMM: v = settings.DefaultVerticalGapMM
        mMatchX(i) = (designGeometry.LeftX + designGeometry.RightX) / 2#
        mMatchY(i) = (designGeometry.BottomY + designGeometry.TopY) / 2#
        If settings.ModelMode <> AA_MODEL_CUSTOM And collision.IsSharpRectangle Then h = 0#: v = 0#
        If settings.HasHorizontalGap Then h = settings.HorizontalGap
        If settings.HasVerticalGap Then v = settings.VerticalGap
        Set box = New AAGeometry
        box.AddBounds designGeometry.LeftX, designGeometry.BottomY, designGeometry.RightX, designGeometry.TopY
        box.AddBounds collision.LeftX, collision.BottomY, collision.RightX, collision.TopY
        Set orientation = New AARotationVariant
        ' Geometri sudah dalam orientasi aktual. Tidak ada pencarian rotasi.
        orientation.Configure 0#, designGeometry, collision, settings.IncludeDesignZones
        Set mCurrent(i) = orientation.MakePlacement(i, -1, 0, box.LeftX, box.TopY, h, v)
        Set mOriginal(i) = mCurrent(i).CopyForPlan
    Next i
End Sub

Private Function Members(ByVal root As Shape) As Collection
    Dim subgroup As Shape, member As Shape, result As New Collection
    If root.Type <> cdrGroupShape Then Fail "Daftarkan group container untuk Optimize."
    ' Tepat dua tingkat: container -> sumber -> instance. Group instance tetap utuh.
    For Each subgroup In root.Shapes
        If subgroup.Type = cdrGroupShape Then
            For Each member In subgroup.Shapes
                result.Add member
            Next member
        Else
            ' AA tidak membuat subgroup sumber untuk satu instance non-group.
            result.Add subgroup
        End If
    Next subgroup
    Set Members = result
End Function

Private Sub AreaOrigin(ByVal root As Shape, ByRef x As Double, ByRef y As Double)
    Dim w As Double, h As Double
    root.GetBoundingBox x, y, w, h, False
    If mSettings.HasCustomRestrictions Then
        ' Zona/margin Custom memakai grid Process tetap, bukan pusat bbox susunan.
        x = mSettings.AreaLeft(mPageCenterX, mSettings.AreaColumn(x + w / 2#, mPageCenterX))
        y = mSettings.AreaBottom(mPageCenterY)
        Exit Sub
    End If
    x = x + w / 2# - mSettings.AreaWidthMM / 2#
    y = y + h / 2# - mSettings.AreaHeightMM / 2#
End Sub

Private Function PositionMatches(ByVal a As Shape, ByVal ax As Double, ByVal ay As Double, _
                                 ByVal b As Shape, ByVal bx As Double, ByVal by As Double) As Boolean
    Dim x As Double, y As Double, w As Double, h As Double, dx As Double, dy As Double
    a.GetBoundingBox x, y, w, h, False
    dx = x + w / 2# - ax: dy = y + h / 2# - ay
    b.GetBoundingBox x, y, w, h, False
    dx = dx - (x + w / 2# - bx): dy = dy - (y + h / 2# - by)
    PositionMatches = (dx * dx + dy * dy <= AA_OPTIMIZE_MATCH_TOLERANCE_MM ^ 2 + 0.000000001)
End Function

Private Function AnglesMatch(ByVal a As Double, ByVal b As Double) As Boolean
    Dim delta As Double
    delta = a - b: delta = delta - 360# * Int((delta + 180#) / 360#)
    AnglesMatch = (Abs(delta) <= AA_OPTIMIZE_ANGLE_TOLERANCE)
End Function

Public Property Get Count() As Long
    Count = mCount
End Property

Public Property Get AreaWidth() As Double
    AreaWidth = mSettings.AreaWidthMM
End Property

Public Property Get AreaHeight() As Double
    AreaHeight = mSettings.AreaHeightMM
End Property

Public Property Get Placement(ByVal index As Long) As AAPlacement
    Set Placement = mCurrent(index)
End Property

Friend Property Get OriginalPlacement(ByVal index As Long) As AAPlacement
    Set OriginalPlacement = mOriginal(index)
End Property

Friend Property Get MatchX(ByVal index As Long) As Double
    MatchX = mMatchX(index)
End Property

Friend Property Get MatchY(ByVal index As Long) As Double
    MatchY = mMatchY(index)
End Property

Friend Property Get MatchAngle(ByVal index As Long) As Double
    MatchAngle = mDesignAngle(index)
End Property

Public Function MatchesTemplate(ByVal other As AAOptimizeLayout) As Boolean
    Dim i As Long, dx As Double, dy As Double
    If mCount <> other.Count Then Exit Function
    For i = 1 To mCount
        dx = mMatchX(i) - other.MatchX(i): dy = mMatchY(i) - other.MatchY(i)
        If dx * dx + dy * dy > AA_OPTIMIZE_MATCH_TOLERANCE_MM ^ 2 + 0.000000001 Then Exit Function
        ' Angle absolut Design/Cut dapat berbeda sejak sumber; bandingkan
        ' orientasi instance Design antarcontainer, bukan dengan artwork Cut.
        If Not AnglesMatch(mDesignAngle(i), other.MatchAngle(i)) Then Exit Function
    Next i
    MatchesTemplate = True
End Function

Public Sub BeginTrial(ByVal index As Long)
    Dim indices As New Collection
    indices.Add index
    BeginGroupTrial indices
End Sub

Public Sub BeginGroupTrial(ByVal indices As Collection)
    Dim i As Long, box As AAGeometry
    ReDim mMoving(1 To mCount)
    For i = 1 To indices.Count
        mMoving(CLng(indices(i))) = True
    Next i
    Set mOccupied = New Collection
    For i = 1 To mCount
        If Not mMoving(i) Then mOccupied.Add mCurrent(i)
        Set box = mCurrent(i).Footprint
        If i = 1 Then
            mEnvelopeLeft = box.LeftX: mEnvelopeBottom = box.BottomY
            mEnvelopeRight = box.RightX: mEnvelopeTop = box.TopY
        Else
            If box.LeftX < mEnvelopeLeft Then mEnvelopeLeft = box.LeftX
            If box.BottomY < mEnvelopeBottom Then mEnvelopeBottom = box.BottomY
            If box.RightX > mEnvelopeRight Then mEnvelopeRight = box.RightX
            If box.TopY > mEnvelopeTop Then mEnvelopeTop = box.TopY
        End If
    Next i
    mIndex.Clear
End Sub

Public Function CanMove(ByVal index As Long, ByVal dx As Double, ByVal dy As Double) As Boolean
    Dim candidate As AAPlacement, other As AAPlacement, obstacle As AAGeometry
    Dim nearby As Collection, box As AAGeometry, gap As Double
    Dim i As Long, oldDX As Double, oldDY As Double, newDX As Double, newDY As Double
    Set candidate = mCurrent(index).CopyForPlan
    If dx <> 0# Or dy <> 0# Then candidate.Translate dx, dy
    Set box = candidate.Footprint
    If box.LeftX < -AA_EDGE_TOLERANCE_MM Or box.BottomY < -AA_EDGE_TOLERANCE_MM Or _
       box.RightX > mSettings.AreaWidthMM + AA_EDGE_TOLERANCE_MM Or _
       box.TopY > mSettings.AreaHeightMM + AA_EDGE_TOLERANCE_MM Then Exit Function
    If box.LeftX < mEnvelopeLeft - AA_EDGE_TOLERANCE_MM Or _
       box.RightX > mEnvelopeRight + AA_EDGE_TOLERANCE_MM Or _
       box.BottomY < mEnvelopeBottom - AA_EDGE_TOLERANCE_MM Or _
       box.TopY > mEnvelopeTop + AA_EDGE_TOLERANCE_MM Then Exit Function
    For Each obstacle In mForbidden
        If candidate.ForbiddenFootprint.Overlaps(obstacle) Then Exit Function
    Next obstacle
    ' Jangan melewati urutan relatif objek semula ketika merapatkan pola.
    For i = 1 To mCount
        If i <> index Then
            oldDX = MidX(mOriginal(index)) - MidX(mOriginal(i))
            oldDY = MidY(mOriginal(index)) - MidY(mOriginal(i))
            newDX = MidX(candidate) - MidX(mCurrent(i))
            newDY = MidY(candidate) - MidY(mCurrent(i))
            If mMoving(i) Then newDX = newDX - dx: newDY = newDY - dy
            If Abs(oldDX) > AA_EDGE_TOLERANCE_MM And oldDX * newDX < 0# Then Exit Function
            If Abs(oldDY) > AA_EDGE_TOLERANCE_MM And oldDY * newDY < 0# Then Exit Function
        End If
    Next i
    Set box = candidate.CollisionFootprint
    gap = candidate.GapH
    If candidate.GapV > gap Then gap = candidate.GapV
    Set nearby = mIndex.Nearby(mOccupied, box.LeftX, box.BottomY, box.RightX, box.TopY, gap)
    For Each other In nearby
        If Not candidate.HasClearance(other) Then Exit Function
        If box.Overlaps(other.CollisionFootprint) Then Exit Function
    Next other
    CanMove = True
End Function

Public Function Valid() As Boolean
    Dim i As Long
    For i = 1 To mCount
        AAProgressPulse
        BeginTrial i
        If Not CanMove(i, 0#, 0#) Then
            AADebugWrite "Optimize", "validation rejected: item=" & CStr(i)
            Exit Function
        End If
    Next i
    Valid = True
End Function

' Minimum memeriksa tetangga yang sudah ditetapkan. Anggota yang masih
' menunggu boleh memiliki gap awal terlalu kecil dan diperbaiki kemudian.
Friend Sub BeginMinimumTrial(ByVal settled As Collection)
    Dim i As Long
    Set mOccupied = New Collection
    For i = 1 To settled.Count
        mOccupied.Add mCurrent(CLng(settled(i)))
    Next i
    mIndex.Clear
End Sub

Friend Function MinimumCanMove(ByVal index As Long, ByVal dx As Double, _
                               ByVal dy As Double, ByVal checkAreaAndMarks As Boolean) As Boolean
    Dim candidate As AAPlacement, other As AAPlacement, obstacle As AAGeometry
    Dim box As AAGeometry, nearby As Collection, gap As Double
    Set candidate = mCurrent(index).CopyForPlan
    If dx <> 0# Or dy <> 0# Then candidate.Translate dx, dy
    If checkAreaAndMarks Then
        Set box = candidate.Footprint
        If box.LeftX < -AA_EDGE_TOLERANCE_MM Or box.BottomY < -AA_EDGE_TOLERANCE_MM Or _
           box.RightX > mSettings.AreaWidthMM + AA_EDGE_TOLERANCE_MM Or _
           box.TopY > mSettings.AreaHeightMM + AA_EDGE_TOLERANCE_MM Then Exit Function
        For Each obstacle In mForbidden
            If candidate.ForbiddenFootprint.Overlaps(obstacle) Then Exit Function
        Next obstacle
    End If
    Set box = candidate.CollisionFootprint
    gap = candidate.GapH
    If candidate.GapV > gap Then gap = candidate.GapV
    Set nearby = mIndex.Nearby(mOccupied, box.LeftX, box.BottomY, box.RightX, box.TopY, gap)
    For Each other In nearby
        If other.DesignIndex <> index Then
            If Not candidate.HasClearance(other) Then Exit Function
            If box.Overlaps(other.CollisionFootprint) Then Exit Function
        End If
    Next other
    MinimumCanMove = True
End Function

Friend Function MinimumValid() As Boolean
    Dim indices As Collection, i As Long
    Set indices = New Collection
    For i = 1 To mCount
        indices.Add i
    Next i
    BeginMinimumTrial indices
    For i = 1 To mCount
        AAProgressPulse
        If Not MinimumCanMove(i, 0#, 0#, True) Then Exit Function
    Next i
    MinimumValid = True
End Function

Public Sub MovePlan(ByVal index As Long, ByVal dx As Double, ByVal dy As Double)
    mCurrent(index).Translate dx, dy
End Sub

Public Sub Apply(ByVal includeCuts As Boolean)
    Dim i As Long, dx As Double, dy As Double, target As Shape
    For i = 1 To mCount
        AAProgressStep i - 1, mCount
        dx = mCurrent(i).Footprint.LeftX - mOriginal(i).Footprint.LeftX
        dy = mCurrent(i).Footprint.TopY - mOriginal(i).Footprint.TopY
        Set target = mDesigns(i)
        If dx <> 0# Or dy <> 0# Then target.Move dx, dy
        If Not AnglesMatch(target.RotationAngle, mDesignAngle(i)) Then Fail "Rotasi Design berubah."
        If includeCuts And mCuts.Count > 0 Then
            Set target = mCuts(i)
            If dx <> 0# Or dy <> 0# Then target.Move dx, dy
            If Not AnglesMatch(target.RotationAngle, mCutAngle(i)) Then Fail "Rotasi Cut Line berubah."
        End If
    Next i
End Sub

Public Sub Restore()
    Dim i As Long, target As Shape, failed As Boolean
    On Error Resume Next
    For i = 1 To mCount
        Err.Clear
        Set target = mDesigns(i)
        target.Move mDesignX(i) - target.LeftX, mDesignY(i) - target.TopY
        If Err.Number <> 0 Then failed = True
        If mCuts.Count > 0 Then
            Err.Clear
            Set target = mCuts(i)
            target.Move mCutX(i) - target.LeftX, mCutY(i) - target.TopY
            If Err.Number <> 0 Then failed = True
        End If
    Next i
    On Error GoTo 0
    If failed Then Fail "Sebagian posisi objek tidak dapat dipulihkan."
End Sub

' Verifikasi ulang kurva aktual sesudah Move; origin area tetap snapshot awal.
Public Function AppliedValid(Optional ByVal minimumAlignment As Boolean = False) As Boolean
    Dim i As Long, design As Shape, cut As Shape
    Dim geometry As AAGeometry, collision As AAGeometry, box As AAGeometry
    Dim orientation As AARotationVariant, h As Double, v As Double
    For i = 1 To mCount
        AAProgressPulse
        Set design = mDesigns(i)
        If Not AnglesMatch(design.RotationAngle, mDesignAngle(i)) Then Exit Function
        Set geometry = New AAGeometry
        If mCuts.Count = 0 Then
            geometry.Capture design
            Set geometry = geometry.OuterBoundary
            Set geometry = geometry.Placed(geometry.LeftX - mDesignOriginX, geometry.TopY - mDesignOriginY)
            Set collision = geometry
        Else
            Set cut = mCuts(i)
            If Not AnglesMatch(cut.RotationAngle, mCutAngle(i)) Then Exit Function
            If mSettings.IncludeDesignZones Then
                geometry.Capture design
                Set geometry = geometry.OuterBoundary
            Else
                geometry.CaptureBounds design
            End If
            Set geometry = geometry.Placed(geometry.LeftX - mDesignOriginX, geometry.TopY - mDesignOriginY)
            Set collision = New AAGeometry
            collision.Capture cut
            Set collision = collision.Placed(collision.LeftX - mCutOriginX, collision.TopY - mCutOriginY)
        End If
        h = mCurrent(i).GapH: v = mCurrent(i).GapV
        Set box = New AAGeometry
        box.AddBounds geometry.LeftX, geometry.BottomY, geometry.RightX, geometry.TopY
        box.AddBounds collision.LeftX, collision.BottomY, collision.RightX, collision.TopY
        Set orientation = New AARotationVariant
        orientation.Configure 0#, geometry, collision, mSettings.IncludeDesignZones
        Set mCurrent(i) = orientation.MakePlacement(i, -1, 0, box.LeftX, box.TopY, h, v)
    Next i
    If minimumAlignment Then
        AppliedValid = MinimumValid
    Else
        AppliedValid = Valid
    End If
End Function

Private Function MidX(ByVal p As AAPlacement) As Double
    MidX = (p.CollisionFootprint.LeftX + p.CollisionFootprint.RightX) / 2#
End Function

Private Function MidY(ByVal p As AAPlacement) As Double
    MidY = (p.CollisionFootprint.TopY + p.CollisionFootprint.BottomY) / 2#
End Function

Private Sub Fail(ByVal message As String)
    Err.Raise vbObjectError + 5360, "AAOptimizeLayout", message
End Sub
