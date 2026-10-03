Option Explicit

Private Const AA_EDGE_TOLERANCE_MM As Double = 0.0001
Private Const AA_MARK_TOLERANCE_MM As Double = 0.2
Private Const AA_SHIFT_STEP_MM As Double = 0.2
Private Const AA_MAX_SHIFT_PROBES_PER_AREA As Long = 50000
Private Const AA_MAX_OBJECTS_PER_SOURCE As Long = 5000
Private Const AA_MAX_CANDIDATES_PER_AREA As Long = 50000
Private Const AA_MAX_CENTERING_PASSES As Long = 20

Public Function BuildMedium(ByRef models() As AAModel, _
                            ByRef designIndexes() As Long, ByVal designCount As Long, _
                            ByRef cutIndexes() As Long, ByVal cutCount As Long, _
                            ByRef quantities() As Long, ByVal settings As AASettings, _
                            ByVal sourceDocument As Document) As AALayoutPlan
    Dim planner As AAMediumPlanner
    Set planner = New AAMediumPlanner
    Set BuildMedium = planner.Build(models, designIndexes, designCount, cutIndexes, _
        cutCount, quantities, settings, sourceDocument, Me)
End Function

Friend Function ForbiddenFor(ByVal settings As AASettings, _
                             ByVal sourceDocument As Document) As Collection
    Set ForbiddenFor = BuildForbidden(settings, sourceDocument)
End Function

Friend Function FinalPlanFits(ByVal plan As AALayoutPlan, _
                              ByVal settings As AASettings, _
                              ByVal sourceDocument As Document) As Boolean
    Dim offsets As Object
    Dim areaIndex As Long

    Set offsets = CreateObject("Scripting.Dictionary")
    For areaIndex = 0 To plan.AreaCount - 1
        offsets.Add CStr(areaIndex), Array(0#, 0#)
    Next areaIndex
    FinalPlanFits = CenteredPlanFits(plan, offsets, _
        BuildForbidden(settings, sourceDocument), settings)
End Function

Friend Function FinalizeMedium(ByVal plan As AALayoutPlan, ByVal settings As AASettings, _
                               ByVal sourceDocument As Document) As Boolean
    Dim offsets As Object, forbidden As Collection
    Dim placement As AAPlacement, shift As Variant
    Dim safe As Boolean
    If plan.DesignCount = 0 Then
        FinalizeMedium = True
        Exit Function
    End If
    Set forbidden = BuildForbidden(settings, sourceDocument)
    Set offsets = CenterOffsets(plan, settings)
    safe = FindSafeOffsets(plan, offsets, forbidden, settings)
    If safe Then safe = CenteredPlanFits(plan, offsets, forbidden, settings)
    For Each placement In plan.Placements
        shift = offsets(CStr(placement.AreaIndex))
        placement.Translate CDbl(shift(0)), CDbl(shift(1))
    Next placement
    ' Saat pergeseran bersama tidak cukup, caller memperbaiki kandidat individual
    ' dari posisi yang sudah dipusatkan, tanpa mengubah Shape dokumen.
    FinalizeMedium = safe
End Function

Friend Function BetterPlan(ByVal candidate As AALayoutPlan, _
                           ByVal reference As AALayoutPlan) As Boolean
    Dim candidateBoundsArea As Double, referenceBoundsArea As Double
    If candidate.AreaCount <> reference.AreaCount Then
        BetterPlan = (candidate.AreaCount < reference.AreaCount)
    ElseIf candidate.DesignCount <> reference.DesignCount Then
        BetterPlan = (candidate.DesignCount > reference.DesignCount)
    Else
        candidateBoundsArea = OccupiedBoundsArea(candidate)
        referenceBoundsArea = OccupiedBoundsArea(reference)
        If candidateBoundsArea <> referenceBoundsArea Then
            BetterPlan = (candidateBoundsArea < referenceBoundsArea)
        Else
            BetterPlan = HasFullerEarlierArea(candidate, reference)
        End If
    End If
End Function

Private Function OccupiedBoundsArea(ByVal plan As AALayoutPlan) As Double
    Dim lefts() As Double, bottoms() As Double
    Dim rights() As Double, tops() As Double
    Dim seen() As Boolean
    Dim placement As AAPlacement, box As AAGeometry
    Dim areaIndex As Long

    If plan.AreaCount = 0 Then Exit Function
    ReDim lefts(0 To plan.AreaCount - 1)
    ReDim bottoms(0 To plan.AreaCount - 1)
    ReDim rights(0 To plan.AreaCount - 1)
    ReDim tops(0 To plan.AreaCount - 1)
    ReDim seen(0 To plan.AreaCount - 1)
    For Each placement In plan.Placements
        areaIndex = placement.AreaIndex
        Set box = placement.CollisionFootprint
        If seen(areaIndex) Then
            If box.LeftX < lefts(areaIndex) Then lefts(areaIndex) = box.LeftX
            If box.BottomY < bottoms(areaIndex) Then bottoms(areaIndex) = box.BottomY
            If box.RightX > rights(areaIndex) Then rights(areaIndex) = box.RightX
            If box.TopY > tops(areaIndex) Then tops(areaIndex) = box.TopY
        Else
            lefts(areaIndex) = box.LeftX: bottoms(areaIndex) = box.BottomY
            rights(areaIndex) = box.RightX: tops(areaIndex) = box.TopY
            seen(areaIndex) = True
        End If
    Next placement
    For areaIndex = 0 To plan.AreaCount - 1
        If seen(areaIndex) Then _
            OccupiedBoundsArea = OccupiedBoundsArea + _
                (rights(areaIndex) - lefts(areaIndex)) * _
                (tops(areaIndex) - bottoms(areaIndex))
    Next areaIndex
End Function

Public Function BuildMinimum(ByRef models() As AAModel, _
                             ByRef designIndexes() As Long, ByVal designCount As Long, _
                             ByRef cutIndexes() As Long, ByVal cutCount As Long, _
                             ByRef quantities() As Long, ByVal settings As AASettings, _
                             ByVal sourceDocument As Document) As AALayoutPlan
    Dim straightPlan As AALayoutPlan
    Dim zigZagPlan As AALayoutPlan
    Dim straightError As String
    Dim zigZagError As String

    If AllActiveRectangles(models, designIndexes, designCount, quantities) Then
        Set BuildMinimum = BuildOne(models, designIndexes, designCount, cutIndexes, _
            cutCount, quantities, settings, sourceDocument, AA_MODE_STRAIGHT)
        Exit Function
    End If

    On Error Resume Next
    Set straightPlan = BuildOne(models, designIndexes, designCount, cutIndexes, _
        cutCount, quantities, settings, sourceDocument, AA_MODE_STRAIGHT)
    If Err.Number <> 0 Then straightError = Err.Description
    Err.Clear
    Set zigZagPlan = BuildOne(models, designIndexes, designCount, cutIndexes, _
        cutCount, quantities, settings, sourceDocument, AA_MODE_ZIGZAG)
    If Err.Number <> 0 Then zigZagError = Err.Description
    On Error GoTo 0

    If straightPlan Is Nothing And zigZagPlan Is Nothing Then Err.Raise _
        vbObjectError + 5230, "AAPlanner", "Straight: " & straightError & _
        vbCrLf & "ZigZag: " & zigZagError
    If straightPlan Is Nothing Then
        Set BuildMinimum = zigZagPlan
    ElseIf zigZagPlan Is Nothing Then
        Set BuildMinimum = straightPlan
    ElseIf zigZagPlan.AreaCount < straightPlan.AreaCount Then
        Set BuildMinimum = zigZagPlan
    ElseIf zigZagPlan.AreaCount = straightPlan.AreaCount And _
        zigZagPlan.DesignCount > straightPlan.DesignCount Then
        Set BuildMinimum = zigZagPlan
    ElseIf zigZagPlan.AreaCount = straightPlan.AreaCount And _
        zigZagPlan.DesignCount = straightPlan.DesignCount Then
        If HasFullerEarlierArea(zigZagPlan, straightPlan) Then
            Set BuildMinimum = zigZagPlan
        Else
            Set BuildMinimum = straightPlan
        End If
    Else
        Set BuildMinimum = straightPlan
    End If
End Function

Private Function AllActiveRectangles(ByRef models() As AAModel, _
                                     ByRef designIndexes() As Long, _
                                     ByVal designCount As Long, _
                                     ByRef quantities() As Long) As Boolean
    Dim i As Long, index As Long
    For i = 0 To designCount - 1
        index = designIndexes(i)
        If quantities(index) <> 0 Then
            If Not models(index).IsRectangle Then Exit Function
        End If
    Next i
    AllActiveRectangles = True
End Function

Private Function HasFullerEarlierArea(ByVal candidate As AALayoutPlan, _
                                     ByVal reference As AALayoutPlan) As Boolean
    Dim areaIndex As Long
    Dim candidateCount As Long
    Dim referenceCount As Long
    For areaIndex = 0 To candidate.AreaCount - 1
        candidateCount = candidate.DesignCountInArea(areaIndex)
        referenceCount = reference.DesignCountInArea(areaIndex)
        If candidateCount <> referenceCount Then
            HasFullerEarlierArea = (candidateCount > referenceCount)
            Exit Function
        End If
    Next areaIndex
End Function

Private Function BuildOne(ByRef models() As AAModel, _
                          ByRef designIndexes() As Long, ByVal designCount As Long, _
                          ByRef cutIndexes() As Long, ByVal cutCount As Long, _
                          ByRef quantities() As Long, ByVal settings As AASettings, _
                          ByVal sourceDocument As Document, _
                          ByVal layoutMode As Long) As AALayoutPlan
    Dim result As AALayoutPlan
    Dim forbidden As Collection
    Dim offsets As Object
    Dim nextOffsets As Object
    Dim placement As AAPlacement
    Dim shift As Variant
    Dim passIndex As Long

    Set forbidden = BuildForbidden(settings, sourceDocument)
    ' Susun tanpa mark dahulu, lalu pusatkan seluruh isi tiap area.
    ' Coba geser seluruh isi area sebelum merencanakan ulang Quantity.
    For passIndex = 1 To AA_MAX_CENTERING_PASSES
        Set result = BuildOneRaw(models, designIndexes, designCount, cutIndexes, _
            cutCount, quantities, settings, layoutMode, forbidden, offsets)
        If result.DesignCount = 0 Then
            Set BuildOne = result
            Exit Function
        End If
        Set nextOffsets = CenterOffsets(result, settings)
        If FindSafeOffsets(result, nextOffsets, forbidden, settings) Then
            If CenteredPlanFits(result, nextOffsets, forbidden, settings) Then
                For Each placement In result.Placements
                    shift = nextOffsets(CStr(placement.AreaIndex))
                    placement.Translate CDbl(shift(0)), CDbl(shift(1))
                Next placement
                Set BuildOne = result
                Exit Function
            End If
        End If
        Set offsets = nextOffsets
    Next passIndex
    Err.Raise vbObjectError + 5234, "AAPlanner", _
        "Pemusatan susunan dan pemeriksaan mark belum menemukan posisi yang stabil."
End Function

Private Function BuildOneRaw(ByRef models() As AAModel, _
                             ByRef designIndexes() As Long, ByVal designCount As Long, _
                             ByRef cutIndexes() As Long, ByVal cutCount As Long, _
                             ByRef quantities() As Long, ByVal settings As AASettings, _
                             ByVal layoutMode As Long, ByVal forbidden As Collection, _
                             ByVal offsets As Object) As AALayoutPlan
    Dim result As AALayoutPlan
    Dim occupied As Collection
    Dim roleIndex As Long
    Dim designIndex As Long
    Dim cutIndex As Long
    Dim areaIndex As Long
    Dim remaining As Long
    Dim added As Long
    Dim sourceCount As Long
    Dim requested As Boolean

    Set result = New AALayoutPlan
    result.Configure layoutMode
    Set occupied = New Collection
    For roleIndex = 0 To designCount - 1
        designIndex = designIndexes(roleIndex)
        cutIndex = -1
        If cutCount = 1 Then
            cutIndex = cutIndexes(0)
        ElseIf cutCount = designCount Then
            cutIndex = cutIndexes(roleIndex)
        End If
        remaining = quantities(designIndex)
        If remaining <> 0 Then requested = True
        If remaining > AA_MAX_OBJECTS_PER_SOURCE Then Err.Raise _
            vbObjectError + 5210, "AAPlanner", _
            "Quantity satu Design melebihi batas 5000 objek."
        If remaining = 0 Then GoTo NextDesign

        If remaining < 0 Then
            added = FillArea(result, models(designIndex), designIndex, cutIndex, _
                areaIndex, occupied, forbidden, settings, layoutMode, _
                AA_MAX_OBJECTS_PER_SOURCE, offsets)
            If added = 0 And occupied.Count > 0 Then
                areaIndex = areaIndex + 1
                Set occupied = New Collection
                added = FillArea(result, models(designIndex), designIndex, cutIndex, _
                    areaIndex, occupied, forbidden, settings, layoutMode, _
                    AA_MAX_OBJECTS_PER_SOURCE, offsets)
            End If
            If added = 0 Then Err.Raise vbObjectError + 5231, "AAPlanner", _
                "Obj. ID sumber tidak dapat muat dalam satu area."
            If cutCount = 1 And result.DesignCount > AA_MAX_OBJECTS_PER_SOURCE Then _
                Err.Raise vbObjectError + 5210, "AAPlanner", _
                "Satu Cut Line bersama melebihi batas 5000 objek."
        Else
            sourceCount = 0
            Do While remaining > 0
                added = FillArea(result, models(designIndex), designIndex, cutIndex, _
                    areaIndex, occupied, forbidden, settings, layoutMode, remaining, offsets)
                If added = 0 And occupied.Count = 0 Then Err.Raise _
                    vbObjectError + 5231, "AAPlanner", _
                    "Obj. ID sumber tidak dapat muat dalam satu area."
                remaining = remaining - added
                sourceCount = sourceCount + added
                If sourceCount > AA_MAX_OBJECTS_PER_SOURCE Then Err.Raise _
                    vbObjectError + 5210, "AAPlanner", _
                    "Satu Design melebihi batas 5000 objek."
                If remaining > 0 Then
                    areaIndex = areaIndex + 1
                    Set occupied = New Collection
                End If
            Loop
            If cutCount = 1 And result.DesignCount > AA_MAX_OBJECTS_PER_SOURCE Then _
                Err.Raise vbObjectError + 5210, "AAPlanner", _
                "Satu Cut Line bersama melebihi batas 5000 objek."
        End If
NextDesign:
    Next roleIndex
    If requested And result.DesignCount = 0 Then Err.Raise _
        vbObjectError + 5233, "AAPlanner", _
        "Tidak ada Design yang muat pada area dan zona mark yang dipilih."
    Set BuildOneRaw = result
End Function

Private Function FillArea(ByVal result As AALayoutPlan, ByVal model As AAModel, _
                          ByVal designIndex As Long, ByVal cutIndex As Long, _
                          ByVal areaIndex As Long, ByVal occupied As Collection, _
                          ByVal forbidden As Collection, ByVal settings As AASettings, _
                          ByVal layoutMode As Long, ByVal limit As Long, _
                          ByVal offsets As Object) As Long
    Dim horizontalStep As Double
    Dim verticalStep As Double
    Dim xOffset As Double
    Dim yOffset As Double
    Dim rowIndex As Long
    Dim probes As Long
    Dim candidate As AAPlacement
    Dim startingRows() As Double
    Dim startIndex As Long
    Dim rowAdded As Boolean
    Dim horizontalGap As Double, verticalGap As Double
    Dim gridTop As Double
    Dim nextTop As Double
    Dim i As Long
    Dim retryBoundary As Boolean

    ' Semua rectangle selalu memakai Straight, termasuk sudut lengkung dan Custom.
    ' layoutMode ByVal: sumber lain tetap mengikuti pola rencana yang dibandingkan.
    If model.IsRectangle Then layoutMode = AA_MODE_STRAIGHT
    model.GetMinimumSteps layoutMode, horizontalStep, verticalStep
    model.GetBoundaryGaps layoutMode, horizontalGap, verticalGap
    BuildStartingRows occupied, designIndex, verticalGap, _
        cutIndex >= 0, model.CollisionInsetY, settings.AreaHeightMM, startingRows
    gridTop = startingRows(0)
    yOffset = settings.AreaHeightMM - gridTop
    Do While yOffset + model.FootprintHeight <= settings.AreaHeightMM + AA_EDGE_TOLERANCE_MM
        rowAdded = False
        xOffset = 0#
        If layoutMode = AA_MODE_ZIGZAG And rowIndex Mod 2 = 1 Then _
            xOffset = horizontalStep / 2#
        Do While xOffset + model.FootprintWidth <= _
                 settings.AreaWidthMM + AA_EDGE_TOLERANCE_MM
            probes = probes + 1
            If probes > AA_MAX_CANDIDATES_PER_AREA Then Err.Raise _
                vbObjectError + 5232, "AAPlanner", _
                "Langkah gap menghasilkan terlalu banyak kandidat posisi."
            Set candidate = model.MakePlacement(designIndex, cutIndex, areaIndex, _
                xOffset + model.DesignInsetX, settings.AreaHeightMM - yOffset - model.DesignInsetY)
            If CandidateFits(candidate, occupied, forbidden, settings, offsets, _
                             horizontalGap, verticalGap) Then
                result.Add candidate
                occupied.Add candidate
                FillArea = FillArea + 1
                rowAdded = True
                If FillArea >= limit Then Exit Function
            End If
            xOffset = xOffset + horizontalStep
        Loop
        If FillArea = 0 And Not rowAdded And startIndex < UBound(startingRows) Then
            startIndex = startIndex + 1
            gridTop = startingRows(startIndex)
        Else
            retryBoundary = False
            nextTop = settings.AreaHeightMM - yOffset - verticalStep
            For i = 0 To UBound(startingRows)
                If startingRows(i) < settings.AreaHeightMM - yOffset - _
                    AA_EDGE_TOLERANCE_MM And startingRows(i) > nextTop + _
                    AA_EDGE_TOLERANCE_MM Then
                    gridTop = startingRows(i)
                    rowIndex = 0
                    retryBoundary = True
                    Exit For
                End If
            Next i
            If Not retryBoundary Then rowIndex = rowIndex + 1
        End If
        yOffset = settings.AreaHeightMM - gridTop + rowIndex * verticalStep
    Loop
End Function

Private Sub BuildStartingRows(ByVal occupied As Collection, ByVal designIndex As Long, _
                              ByVal verticalGap As Double, ByVal hasCut As Boolean, _
                              ByVal collisionInsetY As Double, ByVal areaTop As Double, _
                              ByRef values() As Double)
    Dim unique As Object
    Dim placed As AAPlacement
    Dim key As Variant
    Dim value As Double
    Dim i As Long, j As Long

    Set unique = CreateObject("Scripting.Dictionary")
    unique.Add CStr(areaTop), areaTop
    For Each placed In occupied
        If placed.DesignIndex <> designIndex Then
            value = placed.CollisionFootprint.BottomY - verticalGap + collisionInsetY
            If value >= 0# And value <= areaTop Then
                If Not unique.Exists(CStr(value)) Then unique.Add CStr(value), value
            End If
        End If
    Next placed
    ReDim values(0 To unique.Count - 1)
    For Each key In unique.Keys
        values(i) = CDbl(unique(key))
        i = i + 1
    Next key
    ' Tingkat awal dari atas ke bawah; grid dilanjutkan dari baris pertama yang diterima.
    For i = 1 To UBound(values)
        value = values(i)
        j = i - 1
        Do While j >= 0
            If values(j) >= value Then Exit Do
            values(j + 1) = values(j)
            j = j - 1
        Loop
        values(j + 1) = value
    Next i
End Sub

Private Function CandidateFits(ByVal candidate As AAPlacement, _
                               ByVal occupied As Collection, _
                               ByVal forbidden As Collection, _
                               ByVal settings As AASettings, ByVal offsets As Object, _
                               ByVal horizontalGap As Double, ByVal verticalGap As Double) As Boolean
    Dim placed As AAPlacement
    Dim footprint As AAGeometry
    Dim collisionFootprint As AAGeometry
    Dim spacingFootprint As AAGeometry
    Dim placedSpacing As AAGeometry

    Set footprint = candidate.Footprint
    Set collisionFootprint = candidate.CollisionFootprint
    If footprint.LeftX < -AA_EDGE_TOLERANCE_MM Or _
       footprint.BottomY < -AA_EDGE_TOLERANCE_MM Or _
       footprint.RightX > settings.AreaWidthMM + AA_EDGE_TOLERANCE_MM Or _
       footprint.TopY > settings.AreaHeightMM + AA_EDGE_TOLERANCE_MM Then Exit Function
    If Not offsets Is Nothing Then
        If offsets.Exists(CStr(candidate.AreaIndex)) Then
            If Not PlacementMarksFit(candidate, offsets, forbidden) Then Exit Function
        End If
    End If
    For Each placed In occupied
        ' Acuan pasangan: Cut Line; standalone: siluet Design.
        Set spacingFootprint = collisionFootprint
        Set placedSpacing = placed.CollisionFootprint
        If spacingFootprint.Overlaps(placedSpacing) Then Exit Function
        If placed.DesignIndex <> candidate.DesignIndex Then
            If candidate.UsesContourGap And placed.UsesContourGap Then
                If Not candidate.HasClearance(placed) Then Exit Function
            Else
                ' Gap negatif/persen tetap mengikuti kontrak langkah sebelumnya.
                If spacingFootprint.RightX + horizontalGap > placedSpacing.LeftX + AA_EDGE_TOLERANCE_MM And _
                   spacingFootprint.LeftX - horizontalGap < placedSpacing.RightX - AA_EDGE_TOLERANCE_MM And _
                   spacingFootprint.TopY + verticalGap > placedSpacing.BottomY + AA_EDGE_TOLERANCE_MM And _
                   spacingFootprint.BottomY - verticalGap < placedSpacing.TopY - AA_EDGE_TOLERANCE_MM Then Exit Function
            End If
        End If
    Next placed
    CandidateFits = True
End Function

Private Function CenterOffsets(ByVal plan As AALayoutPlan, _
                               ByVal settings As AASettings) As Object
    Dim bounds As Object
    Dim result As Object
    Dim placement As AAPlacement
    Dim footprint As AAGeometry
    Dim key As Variant
    Dim box As Variant

    Set bounds = CreateObject("Scripting.Dictionary")
    Set result = CreateObject("Scripting.Dictionary")
    For Each placement In plan.Placements
        key = CStr(placement.AreaIndex)
        Set footprint = placement.Footprint
        If bounds.Exists(key) Then
            box = bounds(key)
            If footprint.LeftX < box(0) Then box(0) = footprint.LeftX
            If footprint.BottomY < box(1) Then box(1) = footprint.BottomY
            If footprint.RightX > box(2) Then box(2) = footprint.RightX
            If footprint.TopY > box(3) Then box(3) = footprint.TopY
            bounds(key) = box
        Else
            bounds.Add key, Array(footprint.LeftX, footprint.BottomY, _
                footprint.RightX, footprint.TopY)
        End If
    Next placement
    For Each key In bounds.Keys
        box = bounds(key)
        result.Add key, Array((settings.AreaWidthMM - CDbl(box(0)) - _
            CDbl(box(2))) / 2#, (settings.AreaHeightMM - CDbl(box(1)) - _
            CDbl(box(3))) / 2#)
    Next key
    Set CenterOffsets = result
End Function

Private Function FindSafeOffsets(ByVal plan As AALayoutPlan, ByVal offsets As Object, _
                                 ByVal forbidden As Collection, _
                                 ByVal settings As AASettings) As Boolean
    Dim key As Variant
    Dim placement As AAPlacement
    Dim footprints As Collection
    Dim markFootprints As Collection
    Dim footprint As AAGeometry
    Dim shift As Variant
    Dim dx As Double, dy As Double
    Dim allSafe As Boolean

    allSafe = True
    For Each key In offsets.Keys
        Set footprints = New Collection
        Set markFootprints = New Collection
        For Each placement In plan.Placements
            If CStr(placement.AreaIndex) = CStr(key) Then
                Set footprint = ShiftedFootprint(placement, offsets)
                footprints.Add footprint
                shift = offsets(key)
                Set footprint = placement.CollisionFootprint.Placed( _
                    placement.CollisionFootprint.LeftX + CDbl(shift(0)), _
                    placement.CollisionFootprint.TopY + CDbl(shift(1)))
                markFootprints.Add footprint
            End If
        Next placement
        If FindAreaShift(footprints, markFootprints, forbidden, settings, dx, dy) Then
            shift = offsets(key)
            offsets(key) = Array(CDbl(shift(0)) + dx, CDbl(shift(1)) + dy)
        Else
            allSafe = False
        End If
    Next key
    FindSafeOffsets = allSafe
End Function

Private Function FindAreaShift(ByVal footprints As Collection, _
                               ByVal markFootprints As Collection, _
                               ByVal forbidden As Collection, ByVal settings As AASettings, _
                               ByRef bestX As Double, ByRef bestY As Double) As Boolean
    Dim footprint As AAGeometry
    Dim leftX As Double, bottomY As Double, rightX As Double, topY As Double
    Dim xs() As Double, ys() As Double
    Dim i As Long, j As Long, probes As Long
    Dim distance As Double, bestDistance As Double
    Dim found As Boolean

    bestX = 0#
    bestY = 0#
    If AreaShiftFits(footprints, markFootprints, 0#, 0#, forbidden, settings) Then
        FindAreaShift = True
        Exit Function
    End If
    Set footprint = footprints(1)
    leftX = footprint.LeftX
    bottomY = footprint.BottomY
    rightX = footprint.RightX
    topY = footprint.TopY
    For Each footprint In footprints
        If footprint.LeftX < leftX Then leftX = footprint.LeftX
        If footprint.BottomY < bottomY Then bottomY = footprint.BottomY
        If footprint.RightX > rightX Then rightX = footprint.RightX
        If footprint.TopY > topY Then topY = footprint.TopY
    Next footprint
    ' Kandidat 0,2 mm dan posisi tepat di batas ruang gerak; urut jarak dari pusat.
    BuildShiftValues -leftX, settings.AreaWidthMM - rightX, xs
    BuildShiftValues -bottomY, settings.AreaHeightMM - topY, ys
    bestDistance = settings.AreaWidthMM ^ 2 + settings.AreaHeightMM ^ 2 + 1#
    For i = LBound(xs) To UBound(xs)
        If xs(i) ^ 2 > bestDistance Then Exit For
        For j = LBound(ys) To UBound(ys)
            distance = xs(i) ^ 2 + ys(j) ^ 2
            If distance >= bestDistance Then Exit For
            probes = probes + 1
            If probes > AA_MAX_SHIFT_PROBES_PER_AREA Then GoTo SearchFinished
            If AreaShiftFits(footprints, markFootprints, xs(i), ys(j), forbidden, settings) Then
                found = True
                bestDistance = distance
                bestX = xs(i)
                bestY = ys(j)
            End If
        Next j
    Next i
SearchFinished:
    FindAreaShift = found
End Function

Private Sub BuildShiftValues(ByVal minimum As Double, ByVal maximum As Double, _
                             ByRef values() As Double)
    Dim steps As Double
    Dim stepIndex As Long, count As Long, i As Long, j As Long
    Dim value As Double

    ' Selisih ukuran di dalam AA_EDGE_TOLERANCE_MM tetap boleh memakai offset 0.
    If minimum > 0# Then minimum = 0#
    If maximum < 0# Then maximum = 0#
    steps = Abs(minimum)
    If maximum > steps Then steps = maximum
    steps = Int(steps / AA_SHIFT_STEP_MM)
    If steps > AA_MAX_SHIFT_PROBES_PER_AREA Then steps = AA_MAX_SHIFT_PROBES_PER_AREA
    ReDim values(0 To CLng(steps) * 2 + 2)
    values(0) = 0#
    count = 1
    For stepIndex = 1 To CLng(steps)
        value = stepIndex * AA_SHIFT_STEP_MM
        If -value >= minimum Then
            values(count) = -value
            count = count + 1
        End If
        If value <= maximum Then
            values(count) = value
            count = count + 1
        End If
    Next stepIndex
    ReDim Preserve values(0 To count + 1)
    values(count) = minimum
    values(count + 1) = maximum
    ' Grid sudah urut jarak; sisipkan hanya dua posisi batas yang terakhir.
    For i = count To count + 1
        value = values(i)
        j = i - 1
        Do While j >= 0
            If Abs(values(j)) <= Abs(value) Then Exit Do
            values(j + 1) = values(j)
            j = j - 1
        Loop
        values(j + 1) = value
    Next i
End Sub

Private Function AreaShiftFits(ByVal footprints As Collection, _
                               ByVal markFootprints As Collection, _
                               ByVal dx As Double, ByVal dy As Double, _
                               ByVal forbidden As Collection, _
                               ByVal settings As AASettings) As Boolean
    Dim footprint As AAGeometry

    For Each footprint In footprints
        If footprint.LeftX + dx < -AA_EDGE_TOLERANCE_MM Or _
           footprint.BottomY + dy < -AA_EDGE_TOLERANCE_MM Or _
           footprint.RightX + dx > settings.AreaWidthMM + AA_EDGE_TOLERANCE_MM Or _
           footprint.TopY + dy > settings.AreaHeightMM + AA_EDGE_TOLERANCE_MM Then Exit Function
    Next footprint
    For Each footprint In markFootprints
        If Not MarksFit(footprint, dx, dy, forbidden) Then Exit Function
    Next footprint
    AreaShiftFits = True
End Function

Private Function ShiftedFootprint(ByVal placement As AAPlacement, _
                                  ByVal offsets As Object) As AAGeometry
    Dim shift As Variant
    Dim result As AAGeometry
    Dim footprint As AAGeometry
    shift = offsets(CStr(placement.AreaIndex))
    Set footprint = placement.Footprint
    Set result = New AAGeometry
    ' Pemusatan/batas area hanya memerlukan angka, bukan salinan kurva.
    result.AddBounds footprint.LeftX + CDbl(shift(0)), _
        footprint.BottomY + CDbl(shift(1)), footprint.RightX + CDbl(shift(0)), _
        footprint.TopY + CDbl(shift(1))
    Set ShiftedFootprint = result
End Function

Private Function CenteredPlanFits(ByVal plan As AALayoutPlan, ByVal offsets As Object, _
                                  ByVal forbidden As Collection, _
                                  ByVal settings As AASettings) As Boolean
    Dim placement As AAPlacement
    Dim footprint As AAGeometry
    For Each placement In plan.Placements
        Set footprint = ShiftedFootprint(placement, offsets)
        If footprint.LeftX < -AA_EDGE_TOLERANCE_MM Or _
           footprint.BottomY < -AA_EDGE_TOLERANCE_MM Or _
           footprint.RightX > settings.AreaWidthMM + AA_EDGE_TOLERANCE_MM Or _
           footprint.TopY > settings.AreaHeightMM + AA_EDGE_TOLERANCE_MM Then Exit Function
        If Not PlacementMarksFit(placement, offsets, forbidden) Then Exit Function
    Next placement
    CenteredPlanFits = True
End Function

Private Function PlacementMarksFit(ByVal placement As AAPlacement, _
                                   ByVal offsets As Object, _
                                   ByVal forbidden As Collection) As Boolean
    Dim shift As Variant
    shift = offsets(CStr(placement.AreaIndex))
    PlacementMarksFit = MarksFit(placement.CollisionFootprint, _
        CDbl(shift(0)), CDbl(shift(1)), forbidden)
End Function

Private Function MarksFit(ByVal footprint As AAGeometry, _
                          ByVal dx As Double, ByVal dy As Double, _
                          ByVal forbidden As Collection) As Boolean
    Dim obstacle As AAGeometry
    Dim moved As AAGeometry
    For Each obstacle In forbidden
        ' Salin kurva hanya jika bounding box dekat dengan salah satu mark.
        If footprint.LeftX + dx < obstacle.RightX - AA_EDGE_TOLERANCE_MM And _
           footprint.RightX + dx > obstacle.LeftX + AA_EDGE_TOLERANCE_MM And _
           footprint.BottomY + dy < obstacle.TopY - AA_EDGE_TOLERANCE_MM And _
           footprint.TopY + dy > obstacle.BottomY + AA_EDGE_TOLERANCE_MM Then
            If moved Is Nothing Then Set moved = footprint.Placed( _
                footprint.LeftX + dx, footprint.TopY + dy)
            If moved.Overlaps(obstacle) Then Exit Function
        End If
    Next obstacle
    MarksFit = True
End Function

Private Function BuildForbidden(ByVal settings As AASettings, _
                                ByVal sourceDocument As Document) As Collection
    Dim result As New Collection
    Select Case settings.ModelMode
        Case AA_MODEL_KISS_A
            AddMark result, 6#, 11#, 10#, settings, sourceDocument
            AddMark result, 6#, 474#, 10#, settings, sourceDocument
            AddMark result, 319#, 11#, 10#, settings, sourceDocument
            AddMark result, 319#, 474#, 10#, settings, sourceDocument
            AddMark result, 162.5, 12.5, 12#, settings, sourceDocument
        Case AA_MODEL_DIE_A
            AddMark result, 9.5, 9.5, 15#, settings, sourceDocument
            AddMark result, 9.5, 475.5, 15#, settings, sourceDocument
            AddMark result, 315.5, 9.5, 15#, settings, sourceDocument
            AddMark result, 315.5, 475.5, 15#, settings, sourceDocument
    End Select
    Set BuildForbidden = result
End Function

Private Sub AddMark(ByVal result As Collection, ByVal centerX As Double, _
                    ByVal centerY As Double, ByVal sizeMM As Double, _
                    ByVal settings As AASettings, ByVal sourceDocument As Document)
    Dim mark As AAGeometry
    Dim localX As Double
    Dim localY As Double
    localX = centerX - (settings.PresetPageWidthMM - settings.AreaWidthMM) / 2#
    localY = centerY - (settings.PresetPageHeightMM - settings.AreaHeightMM) / 2#
    Set mark = New AAGeometry
    ' Abaikan penetrasi hingga 0,2 mm dari tepi mark, bukan dari tepi object/area.
    mark.AddRectangle localX - sizeMM / 2# + AA_MARK_TOLERANCE_MM, _
        localY - sizeMM / 2# + AA_MARK_TOLERANCE_MM, _
        localX + sizeMM / 2# - AA_MARK_TOLERANCE_MM, _
        localY + sizeMM / 2# - AA_MARK_TOLERANCE_MM, sourceDocument
    result.Add mark
End Sub
