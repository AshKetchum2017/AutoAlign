Option Explicit

Private Const AA_SHIFT_FACTOR As Double = 0.2
Private Const AA_SHIFT_STEP_MM As Double = 1#
Private Const AA_ZIGZAG_VERTICAL_FACTOR As Double = 0.88
Private Const AA_EDGE_TOLERANCE_MM As Double = 0.0001
Private Const AA_MAX_OBJECTS_PER_SOURCE As Long = 5000
Private Const AA_MAX_PROBES_PER_AREA As Long = 50000
Private Const AA_MAX_SLOTS_PER_AREA As Long = 50000
Private Const AA_MAX_CONTACT_ANCHORS As Long = 16
Private Const AA_MAX_REFINEMENT_PROBES_PER_AREA As Long = 2048
Private Const AA_MAX_REFINEMENT_SLOTS As Long = 2048
Private Const AA_COMPACTION_TOLERANCE_MM As Double = 0.01
Private Const AA_MAX_COMPACTION_PROBES_PER_AREA As Long = 4096
Private Const AA_COMPACTION_PASSES As Long = 2

Public Function Build(ByRef models() As AAModel, ByRef designIndexes() As Long, _
                      ByVal designCount As Long, ByRef cutIndexes() As Long, _
                      ByVal cutCount As Long, ByRef quantities() As Long, _
                      ByVal settings As AASettings, ByVal sourceDocument As Document, _
                      ByVal sharedPlanner As AAPlanner) As AALayoutPlan
    Dim best As AALayoutPlan, minimum As AALayoutPlan, candidate As AALayoutPlan
    Dim forbidden As Collection, angles As Collection, angle As Variant
    Dim firstIndex As Long, i As Long, pattern As Long, lastError As String
    Dim hasZigZag As Boolean
    Set forbidden = sharedPlanner.ForbiddenFor(settings, sourceDocument)
    On Error Resume Next
    Set minimum = sharedPlanner.BuildMinimum(models, designIndexes, designCount, _
        cutIndexes, cutCount, quantities, settings, sourceDocument)
    If Err.Number <> 0 Then lastError = Err.Description
    Err.Clear
    On Error GoTo 0
    If Not minimum Is Nothing Then
        Set candidate = AdaptMinimum(minimum, models)
        If ValidPlan(candidate, forbidden, settings) Then Set best = candidate
    End If
    firstIndex = designIndexes(0)
    For i = 0 To designCount - 1
        If quantities(designIndexes(i)) <> 0 Then
            firstIndex = designIndexes(i)
            Exit For
        End If
    Next i
    For i = 0 To designCount - 1
        If quantities(designIndexes(i)) <> 0 And Not models(designIndexes(i)).IsRectangle Then
            hasZigZag = True
            Exit For
        End If
    Next i
    Set angles = AnglesFor(models(firstIndex), 0#)
    For Each angle In angles
        For pattern = 0 To 1
            If pattern = 1 And Not hasZigZag Then GoTo NextPattern
            Set candidate = Nothing
            On Error Resume Next
            Set candidate = BuildPass(models, designIndexes, designCount, cutIndexes, _
                cutCount, quantities, settings, forbidden, CDbl(angle), (pattern = 1))
            If Err.Number <> 0 Then lastError = Err.Description
            Err.Clear
            On Error GoTo 0
            If Not candidate Is Nothing Then
                If Not sharedPlanner.FinalizeMedium(candidate, settings, sourceDocument) Then
                    Set candidate = RepairCentered(candidate, models, forbidden, settings)
                End If
                If Not candidate Is Nothing Then
                    If best Is Nothing Then
                        Set best = candidate
                    ElseIf sharedPlanner.BetterPlan(candidate, best) Then
                        Set best = candidate
                    End If
                End If
            End If
    NextPattern:
        Next pattern
    Next angle
    If best Is Nothing Then Err.Raise vbObjectError + 5241, "AAMediumPlanner", _
        "Tidak ditemukan susunan Medium yang memenuhi area, gap, dan mark." & _
        IIf(Len(lastError) > 0, vbCrLf & lastError, vbNullString)
    Set candidate = CompactPlan(best, models, forbidden, settings)
    If Not candidate Is Nothing Then
        If Not sharedPlanner.FinalizeMedium(candidate, settings, sourceDocument) Then
            Set candidate = RepairCentered(candidate, models, forbidden, settings)
        End If
        If Not candidate Is Nothing Then
            If ValidPlan(candidate, forbidden, settings) Then Set best = candidate
        End If
    End If
    Set Build = best
End Function

Private Function CompactPlan(ByVal plan As AALayoutPlan, ByRef models() As AAModel, _
                             ByVal forbidden As Collection, _
                             ByVal settings As AASettings) As AALayoutPlan
    Dim current() As AAPlacement, placed As AAPlacement, occupied As Collection
    Dim orientation As AARotationVariant, result As AALayoutPlan
    Dim i As Long, j As Long, pass As Long, probes As Long, areaIndex As Long
    Dim hasCurve As Boolean
    For Each placed In plan.Placements
        If Not models(placed.DesignIndex).IsRectangle Then hasCurve = True
    Next placed
    If Not hasCurve Then Exit Function
    ' Salinan terpisah: jika pemusatan/mark gagal, rencana semula tetap valid.
    ReDim current(1 To plan.DesignCount)
    For i = 1 To plan.DesignCount
        Set placed = plan.Placements(i)
        Set orientation = models(placed.DesignIndex).RotationVariant(placed.RotationDelta)
        Set current(i) = orientation.MakePlacement(placed.DesignIndex, placed.CutIndex, _
            placed.AreaIndex, placed.Footprint.LeftX, placed.Footprint.TopY, placed.GapH, placed.GapV)
        current(i).SetShiftOrigin placed.SlotLeft, placed.SlotTop
    Next i
    For pass = 1 To AA_COMPACTION_PASSES
        areaIndex = -1
        For i = 1 To plan.DesignCount
            If current(i).AreaIndex <> areaIndex Then
                areaIndex = current(i).AreaIndex
                probes = 0
            End If
            ' Rectangle mempertahankan anchor remainder Align Top / Align Left.
            If models(current(i).DesignIndex).IsRectangle Then GoTo NextPlaced
            Set occupied = New Collection
            For j = 1 To plan.DesignCount
                If j <> i And current(j).AreaIndex = areaIndex Then occupied.Add current(j)
            Next j
            Set orientation = models(current(i).DesignIndex).RotationVariant(current(i).RotationDelta)
            Set current(i) = CompactAxis(current(i), orientation, occupied, forbidden, settings, True, probes)
            Set current(i) = CompactAxis(current(i), orientation, occupied, forbidden, settings, False, probes)
NextPlaced:
        Next i
    Next pass
    Set result = New AALayoutPlan
    result.Configure AA_MODE_MEDIUM
    For i = 1 To plan.DesignCount
        result.Add current(i)
    Next i
    Set CompactPlan = result
End Function

Private Function CompactAxis(ByVal placed As AAPlacement, ByVal orientation As AARotationVariant, _
                             ByVal occupied As Collection, ByVal forbidden As Collection, _
                             ByVal settings As AASettings, ByVal horizontal As Boolean, _
                             ByRef probes As Long) As AAPlacement
    Dim maximum As Double, safe As Double, blocked As Double, distance As Double
    Dim candidate As AAPlacement, best As AAPlacement, x As Double, y As Double
    Set best = placed
    ' Tanpa tetangga pada arah gerak, pergeseran hanya memperlebar celah luar.
    If Not HasNeighborInDirection(placed, occupied, horizontal) Then GoTo FinishedCompaction
    If horizontal Then
        maximum = placed.Footprint.LeftX - MaxValue(0#, placed.SlotLeft - _
            orientation.ReferenceWidth * AA_SHIFT_FACTOR)
    Else
        maximum = MinValue(settings.AreaHeightMM, placed.SlotTop + _
            orientation.ReferenceHeight * AA_SHIFT_FACTOR) - placed.Footprint.TopY
    End If
    If maximum <= AA_COMPACTION_TOLERANCE_MM Then GoTo FinishedCompaction
    distance = MinValue(AA_SHIFT_STEP_MM, maximum)
    Do While probes < AA_MAX_COMPACTION_PROBES_PER_AREA
        x = placed.Footprint.LeftX: y = placed.Footprint.TopY
        If horizontal Then x = x - distance Else y = y + distance
        probes = probes + 1
        Set candidate = TryExactSlot(orientation, placed.DesignIndex, placed.CutIndex, _
            placed.AreaIndex, x, y, occupied, forbidden, settings, placed.GapH, placed.GapV)
        If candidate Is Nothing Then
            blocked = distance
            Exit Do
        End If
        candidate.SetShiftOrigin placed.SlotLeft, placed.SlotTop
        Set best = candidate
        safe = distance
        If safe >= maximum Then GoTo FinishedCompaction
        distance = MinValue(safe + AA_SHIFT_STEP_MM, maximum)
    Loop
    ' Merapatkan sampai sisi aman terdekat, tanpa melompati posisi collision.
    Do While blocked - safe > AA_COMPACTION_TOLERANCE_MM And _
        probes < AA_MAX_COMPACTION_PROBES_PER_AREA
        distance = (safe + blocked) / 2#
        x = placed.Footprint.LeftX: y = placed.Footprint.TopY
        If horizontal Then x = x - distance Else y = y + distance
        probes = probes + 1
        Set candidate = TryExactSlot(orientation, placed.DesignIndex, placed.CutIndex, _
            placed.AreaIndex, x, y, occupied, forbidden, settings, placed.GapH, placed.GapV)
        If candidate Is Nothing Then
            blocked = distance
        Else
            candidate.SetShiftOrigin placed.SlotLeft, placed.SlotTop
            Set best = candidate
            safe = distance
        End If
    Loop
FinishedCompaction:
    Set CompactAxis = best
End Function

Private Function HasNeighborInDirection(ByVal placed As AAPlacement, _
                                        ByVal occupied As Collection, _
                                        ByVal horizontal As Boolean) As Boolean
    Dim box As AAGeometry, other As AAGeometry, neighbor As AAPlacement
    Set box = placed.CollisionFootprint
    For Each neighbor In occupied
        Set other = neighbor.CollisionFootprint
        If horizontal Then
            If other.LeftX < box.LeftX - AA_EDGE_TOLERANCE_MM And _
               other.BottomY < box.TopY - AA_EDGE_TOLERANCE_MM And _
               other.TopY > box.BottomY + AA_EDGE_TOLERANCE_MM Then
                HasNeighborInDirection = True
                Exit Function
            End If
        Else
            If other.TopY > box.TopY + AA_EDGE_TOLERANCE_MM And _
               other.LeftX < box.RightX - AA_EDGE_TOLERANCE_MM And _
               other.RightX > box.LeftX + AA_EDGE_TOLERANCE_MM Then
                HasNeighborInDirection = True
                Exit Function
            End If
        End If
    Next neighbor
End Function

Private Function RepairCentered(ByVal plan As AALayoutPlan, ByRef models() As AAModel, _
                                ByVal forbidden As Collection, _
                                ByVal settings As AASettings) As AALayoutPlan
    Dim current() As AAPlacement, occupied As Collection
    Dim orientation As AARotationVariant, candidate As AAPlacement
    Dim result As AALayoutPlan, i As Long, j As Long, probes As Long, areaIndex As Long
    If plan.DesignCount = 0 Then
        Set RepairCentered = plan
        Exit Function
    End If
    ReDim current(1 To plan.DesignCount)
    For i = 1 To plan.DesignCount
        Set current(i) = plan.Placements(i)
    Next i
    areaIndex = -1
    For i = 1 To plan.DesignCount
        If current(i).AreaIndex <> areaIndex Then
            probes = 0
            areaIndex = current(i).AreaIndex
        End If
        Set occupied = New Collection
        For j = 1 To plan.DesignCount
            If j <> i And current(j).AreaIndex = areaIndex Then occupied.Add current(j)
        Next j
        If Not Fits(current(i), occupied, forbidden, settings) Then
            Set orientation = models(current(i).DesignIndex).RotationVariant(current(i).RotationDelta)
            Set candidate = TrySlot(orientation, current(i).DesignIndex, current(i).CutIndex, _
                areaIndex, current(i).SlotLeft, current(i).SlotTop, _
                occupied, forbidden, settings, current(i).GapH, current(i).GapV, _
                probes, AA_MAX_PROBES_PER_AREA)
            If candidate Is Nothing Then Exit Function
            Set current(i) = candidate
        End If
    Next i
    Set result = New AALayoutPlan
    result.Configure AA_MODE_MEDIUM
    For i = 1 To plan.DesignCount
        result.Add current(i)
    Next i
    Set RepairCentered = result
End Function

Private Function AdaptMinimum(ByVal source As AALayoutPlan, _
                              ByRef models() As AAModel) As AALayoutPlan
    Dim result As AALayoutPlan, placed As AAPlacement
    Dim orientation As AARotationVariant
    Dim h As Double, v As Double
    Set result = New AALayoutPlan
    result.Configure AA_MODE_MEDIUM
    For Each placed In source.Placements
        models(placed.DesignIndex).GetBoundaryGaps AA_MODE_STRAIGHT, h, v
        Set orientation = models(placed.DesignIndex).RotationVariant(0#)
        result.Add orientation.MakePlacement(placed.DesignIndex, placed.CutIndex, _
            placed.AreaIndex, placed.Footprint.LeftX, placed.Footprint.TopY, h, v)
    Next placed
    Set AdaptMinimum = result
End Function

Private Function ValidPlan(ByVal plan As AALayoutPlan, ByVal forbidden As Collection, _
                           ByVal settings As AASettings) As Boolean
    Dim occupied As Collection, placed As AAPlacement
    Dim areaIndex As Long
    Set occupied = New Collection
    areaIndex = -1
    For Each placed In plan.Placements
        If placed.AreaIndex <> areaIndex Then
            areaIndex = placed.AreaIndex
            Set occupied = New Collection
        End If
        If Not Fits(placed, occupied, forbidden, settings) Then Exit Function
        occupied.Add placed
    Next placed
    ValidPlan = True
End Function

Private Function BuildPass(ByRef models() As AAModel, ByRef designIndexes() As Long, _
                           ByVal designCount As Long, ByRef cutIndexes() As Long, _
                           ByVal cutCount As Long, ByRef quantities() As Long, _
                           ByVal settings As AASettings, ByVal forbidden As Collection, _
                           ByVal firstAngle As Double, ByVal zigZag As Boolean) As AALayoutPlan
    Dim result As AALayoutPlan, occupied As Collection
    Dim roleIndex As Long, index As Long, cutIndex As Long, areaIndex As Long
    Dim remaining As Long, added As Long, sourceCount As Long, requested As Boolean
    Set result = New AALayoutPlan
    result.Configure AA_MODE_MEDIUM
    Set occupied = New Collection
    For roleIndex = 0 To designCount - 1
        index = designIndexes(roleIndex)
        cutIndex = -1
        If cutCount = 1 Then
            cutIndex = cutIndexes(0)
        ElseIf cutCount = designCount Then
            cutIndex = cutIndexes(roleIndex)
        End If
        remaining = quantities(index)
        If remaining = 0 Then GoTo NextSource
        requested = True
        If remaining > AA_MAX_OBJECTS_PER_SOURCE Then Err.Raise vbObjectError + 5210, _
            "AAMediumPlanner", "Quantity satu Design melebihi batas 5000 objek."
        If remaining < 0 Then
            added = FillArea(result, models(index), index, cutIndex, areaIndex, _
                occupied, forbidden, settings, AA_MAX_OBJECTS_PER_SOURCE, firstAngle, zigZag)
        Else
            sourceCount = 0
            Do While remaining > 0
                added = FillArea(result, models(index), index, cutIndex, areaIndex, _
                    occupied, forbidden, settings, remaining, firstAngle, zigZag)
                If added = 0 And occupied.Count = 0 Then Err.Raise vbObjectError + 5231, _
                    "AAMediumPlanner", "Object tidak muat dalam satu area pada pencarian Medium."
                remaining = remaining - added
                sourceCount = sourceCount + added
                If sourceCount > AA_MAX_OBJECTS_PER_SOURCE Then Err.Raise vbObjectError + 5210, _
                    "AAMediumPlanner", "Satu Design melebihi batas 5000 objek."
                If remaining > 0 Then
                    areaIndex = areaIndex + 1
                    Set occupied = New Collection
                End If
            Loop
        End If
        If cutCount = 1 And result.DesignCount > AA_MAX_OBJECTS_PER_SOURCE Then _
            Err.Raise vbObjectError + 5210, "AAMediumPlanner", _
                "Satu Cut Line bersama melebihi batas 5000 objek."
NextSource:
    Next roleIndex
    If requested And result.DesignCount = 0 Then Err.Raise vbObjectError + 5233, _
        "AAMediumPlanner", "Tidak ada Design yang muat pada pencarian Medium."
    Set BuildPass = result
End Function

Private Function AnglesFor(ByVal model As AAModel, ByVal firstAngle As Double) As Collection
    Dim result As New Collection
    Dim angle As Long, stepAngle As Long
    stepAngle = 45
    If model.IsRectangle Then stepAngle = 90
    If model.IsRectangle And firstAngle / 90# <> Int(firstAngle / 90#) Then firstAngle = 0#
    result.Add firstAngle
    For angle = 0 To 315 Step stepAngle
        If CDbl(angle) <> firstAngle Then result.Add CDbl(angle)
    Next angle
    Set AnglesFor = result
End Function

Private Function FillArea(ByVal result As AALayoutPlan, ByVal model As AAModel, _
                          ByVal designIndex As Long, ByVal cutIndex As Long, _
                          ByVal areaIndex As Long, ByVal occupied As Collection, _
                          ByVal forbidden As Collection, ByVal settings As AASettings, _
                          ByVal limit As Long, ByVal firstAngle As Double, _
                          ByVal zigZag As Boolean) As Long
    Dim angles As Collection, angle As Variant, orientation As AARotationVariant
    Dim slots As Collection, unique As Object, slot As Variant
    Dim candidate As AAPlacement, placed As AAPlacement
    Dim h As Double, v As Double, i As Long, probes As Long, angleBudget As Long, phase As Long
    Dim remainderOnly As Boolean, mainAngle As Double
    model.GetBoundaryGaps AA_MODE_STRAIGHT, h, v
    Set angles = AnglesFor(model, firstAngle)
    mainAngle = CDbl(angles(1))
    angleBudget = AA_MAX_REFINEMENT_PROBES_PER_AREA \ angles.Count
    ' Selesaikan pemindaian slot dasar untuk SEMUA rotasi terlebih dahulu.
    ' Batas probe hanya membatasi penyempurnaan, bukan kapasitas slot dasar.
    For phase = 0 To 1
        For Each angle In angles
            probes = 0
            Set orientation = model.RotationVariant(CDbl(angle))
            Set slots = New Collection
            Set unique = CreateObject("Scripting.Dictionary")
            remainderOnly = model.IsRectangle And CDbl(angle) <> mainAngle And occupied.Count > 0
            If remainderOnly Then
                SeedRectangleRemainder slots, unique, occupied, orientation, settings, h, v, mainAngle
            Else
                SeedGrid slots, unique, orientation, settings, h, v, zigZag And Not model.IsRectangle
                SeedAreaEdges slots, unique, occupied, orientation, settings, h, v, mainAngle
                For Each placed In occupied
                    SeedNeighbors slots, unique, placed, orientation, h, v, phase = 1 And Not model.IsRectangle
                    If phase = 1 And slots.Count >= AA_MAX_REFINEMENT_SLOTS Then Exit For
                Next placed
            End If
            Set slots = SortedSlots(slots)
            i = 1
            Do While i <= slots.Count
                slot = slots(i)
                If phase = 0 Then
                    Set candidate = TryExactSlot(orientation, designIndex, cutIndex, areaIndex, _
                        CDbl(slot(0)), CDbl(slot(1)), occupied, forbidden, settings, h, v)
                Else
                    Set candidate = TrySlot(orientation, designIndex, cutIndex, areaIndex, _
                        CDbl(slot(0)), CDbl(slot(1)), occupied, forbidden, settings, h, v, probes, angleBudget)
                End If
                If Not candidate Is Nothing Then
                    result.Add candidate
                    occupied.Add candidate
                    FillArea = FillArea + 1
                    If FillArea >= limit Then Exit Function
                    If phase = 0 And Not remainderOnly Then _
                        SeedNeighbors slots, unique, candidate, orientation, h, v, False
                End If
                If phase = 1 Then
                    If probes >= angleBudget Or i >= AA_MAX_REFINEMENT_SLOTS Then Exit Do
                ElseIf slots.Count >= AA_MAX_SLOTS_PER_AREA Then
                    Err.Raise vbObjectError + 5242, "AAMediumPlanner", _
                        "Pencarian slot dasar Medium mencapai batas aman; susunan belum selesai."
                End If
                i = i + 1
            Loop
        Next angle
    Next phase
End Function

Private Function TryExactSlot(ByVal orientation As AARotationVariant, ByVal designIndex As Long, _
                              ByVal cutIndex As Long, ByVal areaIndex As Long, _
                              ByVal x As Double, ByVal y As Double, ByVal occupied As Collection, _
                              ByVal forbidden As Collection, ByVal settings As AASettings, _
                              ByVal h As Double, ByVal v As Double) As AAPlacement
    Dim candidate As AAPlacement
    If x < -AA_EDGE_TOLERANCE_MM Or y - orientation.Height < -AA_EDGE_TOLERANCE_MM Or _
       x + orientation.Width > settings.AreaWidthMM + AA_EDGE_TOLERANCE_MM Or _
       y > settings.AreaHeightMM + AA_EDGE_TOLERANCE_MM Then Exit Function
    If Not HasClearanceAt(orientation, x, y, occupied, h, v) Then Exit Function
    Set candidate = orientation.MakePlacement(designIndex, cutIndex, areaIndex, x, y, h, v)
    If Fits(candidate, occupied, forbidden, settings, True) Then Set TryExactSlot = candidate
End Function

Private Function HasClearanceAt(ByVal orientation As AARotationVariant, ByVal x As Double, _
                                ByVal y As Double, ByVal occupied As Collection, _
                                ByVal h As Double, ByVal v As Double) As Boolean
    Dim placed As AAPlacement, other As AAGeometry, gap As Double
    Dim leftX As Double, rightX As Double, topY As Double, bottomY As Double
    leftX = x + orientation.InsetX: rightX = leftX + orientation.ReferenceWidth
    topY = y - orientation.InsetY: bottomY = topY - orientation.ReferenceHeight
    For Each placed In occupied
        Set other = placed.CollisionFootprint
        gap = MaxValue(MaxValue(h, v), MaxValue(placed.GapH, placed.GapV))
        If leftX < other.RightX + gap And rightX > other.LeftX - gap And _
           bottomY < other.TopY + gap And topY > other.BottomY - gap Then
            If Not orientation.HasClearanceAt(x, y, placed, h, v) Then Exit Function
        End If
    Next placed
    HasClearanceAt = True
End Function

Private Sub SeedRectangleRemainder(ByVal slots As Collection, ByVal unique As Object, _
                                   ByVal occupied As Collection, ByVal orientation As AARotationVariant, _
                                   ByVal settings As AASettings, ByVal h As Double, ByVal v As Double, _
                                   ByVal mainAngle As Double)
    Dim placed As AAPlacement, box As AAGeometry
    Dim leftX As Double, rightX As Double, topY As Double, bottomY As Double
    Dim x As Double, y As Double, found As Boolean, pass As Long
    For pass = 0 To 1
        For Each placed In occupied
            If pass = 1 Or Abs(placed.RotationDelta - mainAngle) <= AA_EDGE_TOLERANCE_MM Then
                Set box = placed.Footprint
                If Not found Then
                    leftX = box.LeftX: rightX = box.RightX
                    topY = box.TopY: bottomY = box.BottomY
                    found = True
                Else
                    If box.LeftX < leftX Then leftX = box.LeftX
                    If box.RightX > rightX Then rightX = box.RightX
                    If box.TopY > topY Then topY = box.TopY
                    If box.BottomY < bottomY Then bottomY = box.BottomY
                End If
            End If
        Next placed
        If found Then Exit For
    Next pass
    x = rightX + h
    Do While x <= settings.AreaWidthMM - orientation.Width + orientation.ReferenceWidth * AA_SHIFT_FACTOR
        y = topY
        Do While y >= orientation.Height - orientation.ReferenceHeight * AA_SHIFT_FACTOR
            AddSlot slots, unique, x, y
            If slots.Count >= AA_MAX_SLOTS_PER_AREA Then Exit Sub
            y = y - orientation.ReferenceHeight - v
        Loop
        x = x + orientation.ReferenceWidth + h
    Loop
    y = bottomY - v
    Do While y >= orientation.Height - orientation.ReferenceHeight * AA_SHIFT_FACTOR
        x = leftX
        Do While x <= settings.AreaWidthMM - orientation.Width + orientation.ReferenceWidth * AA_SHIFT_FACTOR
            AddSlot slots, unique, x, y
            If slots.Count >= AA_MAX_SLOTS_PER_AREA Then Exit Sub
            x = x + orientation.ReferenceWidth + h
        Loop
        y = y - orientation.ReferenceHeight - v
    Loop
End Sub

Private Sub SeedGrid(ByVal slots As Collection, ByVal unique As Object, _
                     ByVal orientation As AARotationVariant, ByVal settings As AASettings, _
                     ByVal h As Double, ByVal v As Double, ByVal zigZag As Boolean)
    Dim x As Double, y As Double, stepX As Double, stepY As Double
    Dim rowIndex As Long
    stepX = orientation.ReferenceWidth + h
    stepY = orientation.ReferenceHeight + v
    If zigZag Then stepY = orientation.ReferenceHeight * AA_ZIGZAG_VERTICAL_FACTOR
    If stepX <= 0# Or stepY <= 0# Then Exit Sub
    y = settings.AreaHeightMM
    Do While y >= orientation.Height - orientation.ReferenceHeight * AA_SHIFT_FACTOR
        x = 0#
        If zigZag And rowIndex Mod 2 = 1 Then x = stepX / 2#
        Do While x <= settings.AreaWidthMM - orientation.Width + orientation.ReferenceWidth * AA_SHIFT_FACTOR
            AddSlot slots, unique, x, y
            If slots.Count >= AA_MAX_SLOTS_PER_AREA Then Exit Sub
            x = x + stepX
        Loop
        y = y - stepY
        rowIndex = rowIndex + 1
    Loop
End Sub

Private Sub SeedAreaEdges(ByVal slots As Collection, ByVal unique As Object, _
                          ByVal occupied As Collection, ByVal orientation As AARotationVariant, _
                          ByVal settings As AASettings, ByVal h As Double, ByVal v As Double, _
                          ByVal mainAngle As Double)
    Dim placed As AAPlacement, box As AAGeometry
    Dim leftX As Double, topY As Double, rightX As Double, bottomY As Double
    Dim found As Boolean
    AddSlot slots, unique, settings.AreaWidthMM - orientation.Width, settings.AreaHeightMM
    AddSlot slots, unique, 0#, orientation.Height
    For Each placed In occupied
        If Abs(placed.RotationDelta - mainAngle) <= AA_EDGE_TOLERANCE_MM Then
            Set box = placed.Footprint
            If Not found Then
                leftX = box.LeftX: topY = box.TopY
                rightX = box.RightX: bottomY = box.BottomY
                found = True
            Else
                If box.LeftX < leftX Then leftX = box.LeftX
                If box.TopY > topY Then topY = box.TopY
                If box.RightX > rightX Then rightX = box.RightX
                If box.BottomY < bottomY Then bottomY = box.BottomY
            End If
        End If
    Next placed
    If found Then
        ' Slot baru kanan Align Top / bawah Align Left terhadap susunan utama.
        AddSlot slots, unique, rightX + h, topY
        AddSlot slots, unique, leftX, bottomY - v
    End If
End Sub

Private Sub SeedNeighbors(ByVal slots As Collection, ByVal unique As Object, _
                          ByVal placed As AAPlacement, ByVal orientation As AARotationVariant, _
                          ByVal h As Double, ByVal v As Double, ByVal contourContacts As Boolean)
    Dim box As AAGeometry, a As Variant, b As Variant
    Dim anchors As Collection, ownAnchors As Collection
    Dim i As Long, j As Long, sx As Long, sy As Long
    Dim x As Double, y As Double, gapH As Double, gapV As Double
    Set box = placed.CollisionFootprint
    gapH = h: gapV = v
    If placed.GapH > gapH Then gapH = placed.GapH
    If placed.GapV > gapV Then gapV = placed.GapV
    AddSlot slots, unique, box.RightX + gapH - orientation.InsetX, placed.Footprint.TopY
    AddSlot slots, unique, placed.Footprint.LeftX, box.BottomY - gapV + orientation.InsetY
    AddSlot slots, unique, box.LeftX - gapH - orientation.ReferenceWidth - orientation.InsetX, placed.Footprint.TopY
    If Not contourContacts Then Exit Sub
    Set anchors = placed.Contour.Anchors
    Set ownAnchors = orientation.Contour.Anchors
    For i = 1 To anchors.Count
        If i > AA_MAX_CONTACT_ANCHORS Then Exit For
        a = anchors(i)
        For j = 1 To ownAnchors.Count
            If j > AA_MAX_CONTACT_ANCHORS Then Exit For
            b = ownAnchors(j)
            x = box.LeftX + CDbl(a(0)) - orientation.InsetX - CDbl(b(0))
            y = box.TopY - CDbl(a(1)) + orientation.InsetY + CDbl(b(1))
            For sx = -1 To 1
                For sy = -1 To 1
                    AddSlot slots, unique, x + sx * gapH, y + sy * gapV
                Next sy
            Next sx
            If slots.Count >= AA_MAX_SLOTS_PER_AREA Then Exit Sub
        Next j
    Next i
End Sub

Private Sub AddSlot(ByVal slots As Collection, ByVal unique As Object, _
                    ByVal x As Double, ByVal y As Double)
    Dim key As String
    If slots.Count >= AA_MAX_SLOTS_PER_AREA Then Exit Sub
    key = CStr(Round(x, 4)) & "|" & CStr(Round(y, 4))
    If unique.Exists(key) Then Exit Sub
    unique.Add key, True
    slots.Add Array(x, y)
End Sub

Private Function SortedSlots(ByVal slots As Collection) As Collection
    Dim values() As Variant, i As Long, result As New Collection
    If slots.Count = 0 Then
        Set SortedSlots = result
        Exit Function
    End If
    ReDim values(1 To slots.Count)
    For i = 1 To slots.Count
        values(i) = slots(i)
    Next i
    SortSlots values, 1, UBound(values)
    For i = 1 To UBound(values)
        result.Add values(i)
    Next i
    Set SortedSlots = result
End Function

Private Sub SortSlots(ByRef values() As Variant, ByVal low As Long, ByVal high As Long)
    Dim i As Long, j As Long, pivot As Variant, swap As Variant
    i = low: j = high
    pivot = values((low + high) \ 2)
    Do While i <= j
        Do While BeforeSlot(values(i), pivot)
            i = i + 1
        Loop
        Do While BeforeSlot(pivot, values(j))
            j = j - 1
        Loop
        If i <= j Then
            swap = values(i): values(i) = values(j): values(j) = swap
            i = i + 1: j = j - 1
        End If
    Loop
    If low < j Then SortSlots values, low, j
    If i < high Then SortSlots values, i, high
End Sub

Private Function BeforeSlot(ByVal a As Variant, ByVal b As Variant) As Boolean
    If CDbl(a(1)) <> CDbl(b(1)) Then
        BeforeSlot = (CDbl(a(1)) > CDbl(b(1)))
    Else
        BeforeSlot = (CDbl(a(0)) < CDbl(b(0)))
    End If
End Function

Private Function TrySlot(ByVal orientation As AARotationVariant, ByVal designIndex As Long, _
                         ByVal cutIndex As Long, ByVal areaIndex As Long, _
                         ByVal x As Double, ByVal y As Double, ByVal occupied As Collection, _
                         ByVal forbidden As Collection, ByVal settings As AASettings, _
                         ByVal h As Double, ByVal v As Double, ByRef probes As Long, _
                         ByVal probeLimit As Long) As AAPlacement
    Dim xs() As Double, ys() As Double, i As Long, j As Long
    Dim minX As Double, maxX As Double, minY As Double, maxY As Double
    Dim candidate As AAPlacement, best As AAPlacement
    Dim distance As Double, bestDistance As Double
    Dim obstacle As AAGeometry, placed As AAPlacement
    Dim collisionLeft As Double, collisionTop As Double, gapH As Double, gapV As Double
    minX = MaxValue(-orientation.ReferenceWidth * AA_SHIFT_FACTOR, -x)
    maxX = MinValue(orientation.ReferenceWidth * AA_SHIFT_FACTOR, settings.AreaWidthMM - orientation.Width - x)
    minY = MaxValue(-orientation.ReferenceHeight * AA_SHIFT_FACTOR, orientation.Height - y)
    maxY = MinValue(orientation.ReferenceHeight * AA_SHIFT_FACTOR, settings.AreaHeightMM - y)
    If minX > maxX + AA_EDGE_TOLERANCE_MM Or minY > maxY + AA_EDGE_TOLERANCE_MM Then Exit Function
    ShiftValues minX, maxX, xs
    ShiftValues minY, maxY, ys
    collisionLeft = x + orientation.InsetX
    collisionTop = y - orientation.InsetY
    For Each obstacle In forbidden
        AddShiftValue xs, obstacle.RightX - collisionLeft, minX, maxX
        AddShiftValue xs, obstacle.LeftX - collisionLeft - orientation.ReferenceWidth, minX, maxX
        AddShiftValue ys, obstacle.TopY - collisionTop + orientation.ReferenceHeight, minY, maxY
        AddShiftValue ys, obstacle.BottomY - collisionTop, minY, maxY
    Next obstacle
    For Each placed In occupied
        Set obstacle = placed.CollisionFootprint
        gapH = MaxValue(h, placed.GapH): gapV = MaxValue(v, placed.GapV)
        If obstacle.RightX + gapH >= collisionLeft + minX And _
           obstacle.LeftX - gapH <= collisionLeft + orientation.ReferenceWidth + maxX And _
           obstacle.TopY + gapV >= collisionTop - orientation.ReferenceHeight + minY And _
           obstacle.BottomY - gapV <= collisionTop + maxY Then
            AddShiftValue xs, obstacle.RightX + gapH - collisionLeft, minX, maxX
            AddShiftValue xs, obstacle.LeftX - gapH - collisionLeft - orientation.ReferenceWidth, minX, maxX
            AddShiftValue ys, obstacle.TopY + gapV - collisionTop + orientation.ReferenceHeight, minY, maxY
            AddShiftValue ys, obstacle.BottomY - gapV - collisionTop, minY, maxY
        End If
    Next placed
    bestDistance = 1E+30
    For i = 0 To UBound(xs)
        If xs(i) ^ 2 > bestDistance Then Exit For
        For j = 0 To UBound(ys)
            distance = xs(i) ^ 2 + ys(j) ^ 2
            If distance >= bestDistance Then Exit For
            If probes >= probeLimit Then GoTo Finished
            probes = probes + 1
            If Not HasClearanceAt(orientation, x + xs(i), y + ys(j), occupied, h, v) Then GoTo NextShift
            Set candidate = orientation.MakePlacement(designIndex, cutIndex, areaIndex, _
                x + xs(i), y + ys(j), h, v)
            candidate.SetShiftOrigin x, y
            If Fits(candidate, occupied, forbidden, settings, True) Then
                Set best = candidate
                bestDistance = distance
                If distance = 0# Then GoTo Finished
            End If
NextShift:
        Next j
    Next i
Finished:
    Set TrySlot = best
End Function

Private Sub AddShiftValue(ByRef values() As Double, ByVal value As Double, _
                          ByVal minimum As Double, ByVal maximum As Double)
    Dim i As Long, j As Long
    If value < minimum Or value > maximum Then Exit Sub
    For i = 0 To UBound(values)
        If Abs(values(i) - value) <= AA_EDGE_TOLERANCE_MM Then Exit Sub
    Next i
    i = UBound(values) + 1
    ReDim Preserve values(0 To i)
    j = i - 1
    Do While j >= 0
        If Abs(values(j)) <= Abs(value) Then Exit Do
        values(j + 1) = values(j): j = j - 1
    Loop
    values(j + 1) = value
End Sub

Private Sub ShiftValues(ByVal minimum As Double, ByVal maximum As Double, _
                        ByRef values() As Double)
    Dim count As Long, i As Long, j As Long, value As Double
    ReDim values(0 To CLng((maximum - minimum) / AA_SHIFT_STEP_MM) + 5)
    If minimum <= 0# And maximum >= 0# Then
        values(count) = 0#: count = count + 1
    End If
    value = Int(minimum / AA_SHIFT_STEP_MM) * AA_SHIFT_STEP_MM
    Do While value <= maximum
        If value >= minimum And value <> 0# Then
            values(count) = value: count = count + 1
        End If
        value = value + AA_SHIFT_STEP_MM
    Loop
    values(count) = minimum: count = count + 1
    values(count) = maximum: count = count + 1
    ReDim Preserve values(0 To count - 1)
    For i = 1 To UBound(values)
        value = values(i): j = i - 1
        Do While j >= 0
            If Abs(values(j)) <= Abs(value) Then Exit Do
            values(j + 1) = values(j): j = j - 1
        Loop
        values(j + 1) = value
    Next i
End Sub

Private Function Fits(ByVal candidate As AAPlacement, ByVal occupied As Collection, _
                      ByVal forbidden As Collection, ByVal settings As AASettings, _
                      Optional ByVal clearanceChecked As Boolean = False) As Boolean
    Dim box As AAGeometry, collision As AAGeometry, other As AAGeometry
    Dim placed As AAPlacement, obstacle As AAGeometry, gap As Double
    Set box = candidate.Footprint
    If box.LeftX < -AA_EDGE_TOLERANCE_MM Or box.BottomY < -AA_EDGE_TOLERANCE_MM Or _
       box.RightX > settings.AreaWidthMM + AA_EDGE_TOLERANCE_MM Or _
       box.TopY > settings.AreaHeightMM + AA_EDGE_TOLERANCE_MM Then Exit Function
    Set collision = candidate.CollisionFootprint
    For Each obstacle In forbidden
        If collision.Overlaps(obstacle) Then Exit Function
    Next obstacle
    For Each placed In occupied
        Set other = placed.CollisionFootprint
        gap = MaxValue(MaxValue(candidate.GapH, candidate.GapV), MaxValue(placed.GapH, placed.GapV))
        If collision.LeftX < other.RightX + gap And collision.RightX > other.LeftX - gap And _
           collision.BottomY < other.TopY + gap And collision.TopY > other.BottomY - gap Then
            If Not clearanceChecked Then
                If Not candidate.HasClearance(placed) Then Exit Function
            End If
            If collision.Overlaps(other) Then Exit Function
        End If
    Next placed
    Fits = True
End Function

Private Function MinValue(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinValue = a Else MinValue = b
End Function

Private Function MaxValue(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxValue = a Else MaxValue = b
End Function
