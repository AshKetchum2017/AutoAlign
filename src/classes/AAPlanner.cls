Option Explicit

Private Const AA_EDGE_TOLERANCE_MM As Double = 0.0001
Private Const AA_PLAN_BOUNDS_TOLERANCE_MM As Double = 0.001
Private Const AA_MARK_TOLERANCE_MM As Double = 0.2
Private Const AA_SHIFT_STEP_MM As Double = 0.2
Private Const AA_MAX_SHIFT_PROBES_PER_AREA As Long = 50000
Private Const AA_MAX_SHIFT_CACHE_ENTRIES As Long = 32768
Private Const AA_MAX_OBJECTS_PER_SOURCE As Long = 5000
Private Const AA_MAX_CANDIDATES_PER_AREA As Long = 50000
Private Const AA_MAX_CENTERING_PASSES As Long = 20
Private Const AA_MAX_MINIMUM_NEIGHBOR_PROBES As Long = 256
Private Const AA_MAX_MINIMUM_ANCHORS As Long = 64
Private Const AA_MINIMUM_NEIGHBOR_RESERVE_MM As Double = 0.001
Private Const AA_MAX_MINIMUM_MARK_SLOTS As Long = 8
Private Const AA_MAX_MINIMUM_MARK_PROBES As Long = 256
Private Const AA_MINIMUM_MARK_FINE_STEP_MM As Double = 0.025
Private Const AA_MINIMUM_MARK_FINE_STEPS As Long = 10
Private Const AA_MAX_MINIMUM_REFILL_PROBES As Long = 1024
Private Const AA_MAX_MINIMUM_REFILL_TARGETS As Long = 16
Private Const AA_MINIMUM_REFILL_ALIGNMENT_MM As Double = 0.01
Private Const AA_GROUP_REFILL_PROBES As Long = 2048
Private Const AA_GROUP_REFILL_ADDITIONS As Long = 16
Private Const AA_GROUP_LOCAL_SLOTS As Long = 32
Private Const AA_GROUP_LOCAL_STEPS As Long = 16
Private Const AA_GROUP_LOCAL_RADIUS As Double = 0.2
Private Const AA_KISS_CORNER_DIAMETER_MM As Double = 7#
Private Const AA_KISS_CORNER_FALLBACK_MM As Double = 10#
Private Const AA_KISS_ANCHOR_SEARCH_MM As Double = 10#
Private Const AA_KISS_ANCHOR_TOLERANCE_MM As Double = 0.000001

Public Function BuildMedium(ByRef models() As AAModel, _
                            ByRef designIndexes() As Long, ByVal designCount As Long, _
                            ByRef cutIndexes() As Long, ByVal cutCount As Long, _
                            ByRef quantities() As Long, ByVal settings As AASettings, _
                            ByVal sourceDocument As Document, _
                            Optional ByVal singleAreaOnly As Boolean = False, _
                            Optional ByVal capacityTarget As Long = -1) As AALayoutPlan
    Dim planner As AAMediumPlanner
    Set planner = New AAMediumPlanner
    Set BuildMedium = planner.Build(models, designIndexes, designCount, cutIndexes, _
        cutCount, quantities, settings, sourceDocument, Me, singleAreaOnly, capacityTarget)
End Function

' Perbaikan lokal setelah quota final: maksimal satu filler tambahan, tanpa
' mengubah sumber lain atau rotasi. Tidak dijalankan di setiap trial quota.
Friend Function RefillMinimum(ByVal seed As AALayoutPlan, ByVal model As AAModel, _
                              ByVal designIndex As Long, ByVal cutIndex As Long, ByVal maximum As Long, _
                              ByVal settings As AASettings, ByVal sourceDocument As Document) As AALayoutPlan
    Dim original As AAPlacement, neighbor As AAPlacement, placed As AAPlacement, moved As AAPlacement
    Dim occupied As Collection, trialOccupied As Collection, forbidden As Collection, offsets As Object
    Dim candidate As AALayoutPlan, best As AALayoutPlan, count As Long, probes As Long, site As Long, targets As Long
    Dim bestScore As Long, bestDX As Double, bestDY As Double
    Dim sx As Double, sy As Double, zx As Double, zy As Double, h As Double, v As Double
    Dim x As Double, y As Double, dx As Double, dy As Double, started As Double
    Set RefillMinimum = seed: Set best = seed: bestScore = -2147483647
    If seed.AreaCount <> 1 Then Exit Function
    For Each placed In seed.placements
        If placed.designIndex = designIndex Then
            count = count + 1
            If placed.RotationDelta <> 0# Then Exit Function
        End If
    Next placed
    If count = 0 Or count >= maximum Or maximum > AA_MAX_OBJECTS_PER_SOURCE Then Exit Function
    started = AADebugPerfBegin()
    model.GetMinimumSteps AA_MODE_STRAIGHT, sx, sy
    model.GetMinimumSteps AA_MODE_ZIGZAG, zx, zy
    model.GetBoundaryGaps AA_MODE_STRAIGHT, h, v
    If sx <= AA_EDGE_TOLERANCE_MM Or sy <= AA_EDGE_TOLERANCE_MM Or zy <= AA_EDGE_TOLERANCE_MM Then Exit Function
    Set forbidden = ForbiddenFor(settings, sourceDocument)
    If settings.modelMode = AA_MODEL_KISS_A Then Set forbidden = NominalForbiddenFor(settings, sourceDocument)
    Set offsets = CreateObject("Scripting.Dictionary"): offsets.Add "0", Array(0#, 0#)
    For Each original In seed.placements
        AAProgressPulse
        If original.designIndex = designIndex Then
            If targets >= AA_MAX_MINIMUM_REFILL_TARGETS Then GoTo Finished
            targets = targets + 1
            dx = 0#: dy = 0#
            ' Tambahan langsung tetap kandidat; jangan berhenti sebelum
            ' membandingkan kerapian alternatif dengan jumlah yang sama.
            Set candidate = MinimumFillerBonus(seed, original, original, seed.placements, model, _
                designIndex, cutIndex, forbidden, offsets, settings, sourceDocument, sx, sy, zy, h, v, probes, bestScore)
            If Not candidate Is Nothing Then
                Set best = candidate: bestDX = 0#: bestDY = 0#
            End If
            Set occupied = New Collection
            For Each placed In seed.placements
                If Not placed Is original Then occupied.Add placed
            Next placed
            For Each neighbor In occupied
                AAProgressPulse
                If neighbor.designIndex = designIndex Then
                    If Abs(neighbor.footprint.leftX - original.footprint.leftX) <= 2# * sx And _
                       Abs(neighbor.footprint.topY - original.footprint.topY) <= 2# * sy Then
                        For site = 0 To 7
                            If probes >= AA_MAX_MINIMUM_REFILL_PROBES Then GoTo Finished
                            MinimumLocalSlot neighbor.footprint, site, sx, sy, zy, x, y
                            dx = x - original.footprint.leftX: dy = y - original.footprint.topY
                            If Abs(dx) <= sx + AA_EDGE_TOLERANCE_MM And Abs(dy) <= sy + AA_EDGE_TOLERANCE_MM And _
                               Abs(dx) + Abs(dy) > AA_EDGE_TOLERANCE_MM Then
                                probes = probes + 1: AAProgressPulse
                                Set moved = original.CopyForPlan(): moved.Translate dx, dy
                                If CandidateFits(moved, occupied, forbidden, settings, offsets, h, v) Then
                                    Set trialOccupied = New Collection
                                    For Each placed In occupied
                                        trialOccupied.Add placed
                                    Next placed
                                    trialOccupied.Add moved
                                    Set candidate = MinimumFillerBonus(seed, original, moved, trialOccupied, model, _
                                        designIndex, cutIndex, forbidden, offsets, settings, sourceDocument, sx, sy, zy, h, v, probes, bestScore)
                                    If Not candidate Is Nothing Then
                                        Set best = candidate: bestDX = dx: bestDY = dy
                                    End If
                                End If
                            End If
                        Next site
                    End If
                End If
            Next neighbor
        End If
        If probes >= AA_MAX_MINIMUM_REFILL_PROBES Then GoTo Finished
    Next original
Finished:
    Set RefillMinimum = best
    If best.designCount > seed.designCount Then AADebugWrite "Minimum filler refill", _
        "move X/Y=" & Format$(bestDX, "0.000000") & "/" & Format$(bestDY, "0.000000") & _
        "; alignment gain=" & CStr(bestScore)
    AADebugStageEnd "Minimum filler local refill", started
    AADebugWrite "Minimum filler refill", "design=" & CStr(designIndex) & "; probes=" & CStr(probes) & _
        "; targets=" & CStr(targets) & "; added=" & CStr(best.designCount - seed.designCount)
End Function

Private Function MinimumFillerBonus(ByVal seed As AALayoutPlan, ByVal original As AAPlacement, _
                                    ByVal moved As AAPlacement, ByVal occupied As Collection, ByVal model As AAModel, _
                                    ByVal designIndex As Long, ByVal cutIndex As Long, ByVal forbidden As Collection, _
                                    ByVal offsets As Object, ByVal settings As AASettings, ByVal sourceDocument As Document, _
                                    ByVal sx As Double, ByVal sy As Double, ByVal zy As Double, _
                                    ByVal h As Double, ByVal v As Double, ByRef probes As Long, _
                                    ByRef bestScore As Long) As AALayoutPlan
    Dim bonus As AAPlacement, placed As AAPlacement, result As AALayoutPlan
    Dim site As Long, x As Double, y As Double, score As Long
    For site = 0 To 7
        If probes >= AA_MAX_MINIMUM_REFILL_PROBES Then Exit Function
        MinimumLocalSlot moved.footprint, site, sx, sy, zy, x, y
        probes = probes + 1: AAProgressPulse
        Set bonus = model.MakePlacement(designIndex, cutIndex, 0, x + model.DesignInsetX, y - model.DesignInsetY)
        If CandidateFits(bonus, occupied, forbidden, settings, offsets, h, v) Then
            score = MinimumFillerAlignment(seed, original, moved, bonus)
            If score > bestScore Then
                Set result = New AALayoutPlan: result.Configure seed.layoutMode: result.SetRegularPattern False
                For Each placed In seed.placements
                    AAProgressPulse
                    If placed Is original Then
                        result.Add moved.CopyForPlan()
                    Else
                        result.Add placed.CopyForPlan()
                    End If
                Next placed
                result.Add bonus
                If FinalPlanFits(result, settings, sourceDocument) Then
                    Set MinimumFillerBonus = result: bestScore = score
                End If
            End If
        End If
    Next site
End Function

Private Function MinimumFillerAlignment(ByVal seed As AALayoutPlan, ByVal original As AAPlacement, _
                                       ByVal moved As AAPlacement, ByVal bonus As AAPlacement) As Long
    Dim placed As AAPlacement, score As Long
    score = MinimumAxisAlignment(moved.footprint, bonus.footprint)
    For Each placed In seed.placements
        If placed.designIndex = original.designIndex And Not placed Is original Then
            ' Delta seluruh pola: kurangi hubungan lama yang hilang, bukan
            ' hanya menghitung hubungan dua objek kandidat secara terpisah.
            score = score + MinimumAxisAlignment(moved.footprint, placed.footprint) + _
                MinimumAxisAlignment(bonus.footprint, placed.footprint) - _
                MinimumAxisAlignment(original.footprint, placed.footprint)
        End If
    Next placed
    MinimumFillerAlignment = score
End Function

Private Function MinimumAxisAlignment(ByVal first As AAGeometry, ByVal second As AAGeometry) As Long
    Dim score As Long
    If Abs(first.leftX - second.leftX) <= AA_MINIMUM_REFILL_ALIGNMENT_MM Then score = score + 1
    If Abs(first.topY - second.topY) <= AA_MINIMUM_REFILL_ALIGNMENT_MM Then score = score + 1
    MinimumAxisAlignment = score
End Function

Private Sub MinimumLocalSlot(ByVal anchor As AAGeometry, ByVal site As Long, ByVal sx As Double, _
                             ByVal sy As Double, ByVal zy As Double, ByRef x As Double, ByRef y As Double)
    x = anchor.leftX: y = anchor.topY
    Select Case site
        Case 0: x = x + sx
        Case 1: x = x - sx
        Case 2: y = y + sy
        Case 3: y = y - sy
        Case 4: x = x + sx / 2#: y = y + zy
        Case 5: x = x - sx / 2#: y = y + zy
        Case 6: x = x + sx / 2#: y = y - zy
        Case 7: x = x - sx / 2#: y = y - zy
    End Select
End Sub

' Medium membandingkan keluarga blok 0 derajat di luar trial kapasitas.
' Preset yang tidak mengizinkan 0 tetap memakai hasil planner Medium semula.
Friend Function ImproveMediumGroup(ByVal seed As AALayoutPlan, ByRef models() As AAModel, _
                                   ByRef indexes() As Long, ByVal count As Long, _
                                   ByRef cuts() As Long, ByVal cutCount As Long, _
                                   ByVal difference As Long, ByVal settings As AASettings, _
                                   ByVal doc As Document) As AALayoutPlan
    Dim i As Long, candidate As AALayoutPlan, proposed As AALayoutPlan, adapter As AAMediumPlanner
    Set ImproveMediumGroup = seed
    For i = 0 To count - 1
        AAProgressPulse
        If Not models(indexes(i)).AllowsRotation(0#) Then
            AADebugWrite "Medium blocks", "skipped: zero rotation is not allowed for all Designs"
            Exit Function
        End If
    Next i
    Set proposed = ImproveMinimumGroup(seed, models, indexes, count, cuts, cutCount, difference, settings, doc)
    If proposed Is seed Then Exit Function
    ' Gunakan metadata placement Medium agar sumber/Cut yang dipakai ulang
    ' kembali ke sudut sumber, termasuk setelah group sebelumnya diputar.
    Set adapter = New AAMediumPlanner
    Set candidate = adapter.AdaptMinimum(proposed, models)
    candidate.SetRegularPattern proposed.HasRegularPattern
    For i = 1 To proposed.DesignCount
        If Not proposed.Placements(i).CreatesCutLine Then candidate.Placements(i).SuppressCutLine
    Next i
    If Not MinimumGroupValid(candidate, settings, doc) Then
        AADebugWrite "Medium blocks", "rejected after Medium placement adaptation"
        Exit Function
    End If
    AADebugWrite "Medium blocks", "zero-angle candidate; before=" & CStr(seed.DesignCount) & _
        "; selected=" & CStr(candidate.DesignCount)
    Set ImproveMediumGroup = candidate
End Function

' Keluarga blok Minimum untuk kelompok /n bebas; juga kandidat tambahan Medium.
' Tidak dipanggil oleh trial kuota.
Friend Function ImproveMinimumGroup(ByVal seed As AALayoutPlan, ByRef models() As AAModel, _
                                    ByRef indexes() As Long, ByVal count As Long, _
                                    ByRef cuts() As Long, ByVal cutCount As Long, _
                                    ByVal difference As Long, ByVal settings As AASettings, _
                                    ByVal doc As Document) As AALayoutPlan
    Dim best As AALayoutPlan, candidate As AALayoutPlan, allocation() As Long
    Dim placed As AAPlacement, pattern As Long, bestPattern As Long, score As Double, bestScore As Double
    Dim started As Double, orderer As AAMediumPlanner, totalLimit As Long, zeroSeed As Boolean
    Dim i As Long, cut As Long, previousCut As Long, hasSharedCut As Boolean, blockMode As Long, joinGrid As Boolean
    Set ImproveMinimumGroup = seed
    If seed.AreaCount <> 1 Or count < 2 Or difference < 0 Then Exit Function
    zeroSeed = True
    ReDim allocation(0 To UBound(models))
    For Each placed In seed.Placements
        If placed.RotationDelta <> 0# Then zeroSeed = False
        allocation(placed.DesignIndex) = allocation(placed.DesignIndex) + 1
    Next placed
    previousCut = -1
    For i = 0 To count - 1
        If allocation(indexes(i)) > 0 Then
            cut = -1
            If cutCount = 1 Then cut = cuts(0)
            If cutCount > 1 Then cut = cuts(i)
            If cut >= 0 And cut = previousCut Then hasSharedCut = True
            previousCut = cut
        End If
    Next i
    totalLimit = 2147483647
    If cutCount = 1 Then totalLimit = AA_MAX_OBJECTS_PER_SOURCE
    started = AADebugPerfBegin()
    Set best = seed: bestScore = MinimumGroupAlignment(seed)
    For pattern = AA_MODE_STRAIGHT To AA_MODE_ZIGZAG + 3
        AAProgressPulse
        Set candidate = Nothing
        blockMode = pattern: joinGrid = True
        If pattern <= AA_MODE_ZIGZAG Then
            Set candidate = MinimumBlockSeed(models, indexes, count, cuts, cutCount, _
                allocation, blockMode, settings, doc, seed, joinGrid)
        ElseIf pattern = AA_MODE_ZIGZAG + 1 Then
            ' Perbaikan lokal pada hasil lama juga dibandingkan; seed tidak dimutasi.
            If zeroSeed Then Set candidate = seed
        ElseIf hasSharedCut Then
            ' Pertahankan kandidat band terpisah sebagai pembanding jumlah.
            blockMode = pattern - 3: joinGrid = False
            Set candidate = MinimumBlockSeed(models, indexes, count, cuts, cutCount, _
                allocation, blockMode, settings, doc, seed, joinGrid)
        End If
        If Not candidate Is Nothing Then
            Set candidate = MinimumGroupRefill(candidate, models, indexes, count, difference, totalLimit, settings, doc)
            score = MinimumGroupAlignment(candidate)
            If candidate.DesignCount > best.DesignCount Or _
               (candidate.DesignCount = best.DesignCount And score > bestScore) Then
                Set best = candidate: bestScore = score: bestPattern = pattern
            ElseIf candidate.DesignCount = best.DesignCount And score = bestScore Then
                If BetterPlan(candidate, best) Then Set best = candidate: bestPattern = pattern
            End If
            AADebugWrite "Minimum blocks", "pattern=" & CStr(pattern) & "; objects=" & _
                CStr(candidate.DesignCount) & "; alignment=" & CStr(score)
        Else
            AADebugWrite "Minimum blocks", "pattern=" & CStr(pattern) & "; unavailable"
        End If
    Next pattern
    If Not best Is seed Then
        Set orderer = New AAMediumPlanner
        Set best = orderer.OrderRefilledPattern(best)
    End If
    AADebugWrite "Minimum blocks", "before=" & CStr(seed.DesignCount) & "; selected=" & CStr(best.DesignCount) & _
        "; selected pattern=" & CStr(bestPattern)
    AADebugStageEnd "Minimum group blocks + refill", started
    Set ImproveMinimumGroup = best
End Function

Private Function MinimumBlockSeed(ByRef models() As AAModel, ByRef indexes() As Long, _
                                  ByVal count As Long, ByRef cuts() As Long, ByVal cutCount As Long, _
                                  ByRef allocation() As Long, ByVal pattern As Long, _
                                  ByVal settings As AASettings, ByVal doc As Document, _
                                  ByVal reference As AALayoutPlan, ByVal continueSharedCut As Boolean) As AALayoutPlan
    Dim result As AALayoutPlan, occupied As New Collection, forbidden As New Collection
    Dim placed As AAPlacement, model As AAModel, i As Long, index As Long, cut As Long
    Dim row As Long, added As Long, probes As Long, mode As Long, hasBlock As Boolean
    Dim topY As Double, bottomY As Double, x As Double, y As Double, sx As Double, sy As Double
    Dim h As Double, v As Double, previousV As Double
    Dim previousH As Double, previousSX As Double, previousSY As Double
    Dim previousIndex As Long, previousCut As Long, previousMode As Long, sameGrid As Boolean
    Dim cutPolicies As Object, policyCounts() As Long, key As String
    Set result = New AALayoutPlan: result.Configure pattern
    Set cutPolicies = CreateObject("Scripting.Dictionary")
    ReDim policyCounts(0 To UBound(models))
    If Not reference Is Nothing Then
        For Each placed In reference.Placements
            AAProgressPulse
            index = placed.DesignIndex: policyCounts(index) = policyCounts(index) + 1
            cutPolicies.Add CStr(index) & "|" & CStr(policyCounts(index)), placed.CreatesCutLine
        Next placed
    End If
    topY = settings.AreaHeightMM
    For i = 0 To count - 1
        AAProgressPulse
        index = indexes(i): Set model = models(index)
        If allocation(index) > 0 Then
            cut = -1
            If cutCount = 1 Then cut = cuts(0)
            If cutCount > 1 Then cut = cuts(i)
            mode = pattern
            If model.IsRectangle Then mode = AA_MODE_STRAIGHT
            model.GetMinimumSteps mode, sx, sy
            model.GetBoundaryGaps AA_MODE_STRAIGHT, h, v
            If sx <= AA_EDGE_TOLERANCE_MM Or sy <= AA_EDGE_TOLERANCE_MM Then Exit Function
            sameGrid = False
            If continueSharedCut And hasBlock And cut >= 0 And cut = previousCut And mode = previousMode Then
                If Abs(sx - previousSX) <= AA_EDGE_TOLERANCE_MM And _
                   Abs(sy - previousSY) <= AA_EDGE_TOLERANCE_MM And _
                   Abs(h - previousH) <= AA_EDGE_TOLERANCE_MM And _
                   Abs(v - previousV) <= AA_EDGE_TOLERANCE_MM And _
                   Abs(model.FootprintWidth - models(previousIndex).FootprintWidth) <= AA_EDGE_TOLERANCE_MM And _
                   Abs(model.FootprintHeight - models(previousIndex).FootprintHeight) <= AA_EDGE_TOLERANCE_MM Then _
                    sameGrid = models(previousIndex).CanShareCutWith(model)
            End If
            If sameGrid Then
                AADebugWrite "Group block grid", "design=" & CStr(index) & "; continues design=" & _
                    CStr(previousIndex) & "; Cut=" & CStr(cut)
            Else
                If hasBlock Then
                    If previousV > v Then topY = bottomY - previousV Else topY = bottomY - v
                End If
                row = 0: bottomY = topY: y = topY: x = 0#
            End If
            added = 0
            Do While y - model.FootprintHeight >= -AA_EDGE_TOLERANCE_MM And added < allocation(index)
                Do While x + model.FootprintWidth <= settings.AreaWidthMM + AA_EDGE_TOLERANCE_MM _
                         And added < allocation(index)
                    AAProgressPulse
                    probes = probes + 1
                    If probes > AA_MAX_CANDIDATES_PER_AREA Then Exit Function
                    Set placed = model.MakePlacement(index, cut, 0, x + model.DesignInsetX, y - model.DesignInsetY)
                    If CandidateFits(placed, occupied, forbidden, settings, Nothing, h, v) Then
                        key = CStr(index) & "|" & CStr(added + 1)
                        If cutPolicies.Exists(key) Then
                            If Not CBool(cutPolicies(key)) Then placed.SuppressCutLine
                        End If
                        result.Add placed: occupied.Add placed: added = added + 1
                        If placed.Footprint.BottomY < bottomY Then bottomY = placed.Footprint.BottomY
                    End If
                    x = x + sx
                Loop
                If added < allocation(index) Then
                    row = row + 1: y = topY - row * sy: x = 0#
                    If mode = AA_MODE_ZIGZAG And row Mod 2 = 1 Then x = sx / 2#
                End If
            Loop
            If added <> allocation(index) Then Exit Function
            previousV = v: hasBlock = True
            previousH = h: previousSX = sx: previousSY = sy
            previousIndex = index: previousCut = cut: previousMode = mode
        End If
    Next i
    ' Tiap sumber berada di band kotaknya sendiri. Setelah itu hanya kontur
    ' asli yang menjadi obstacle, dan common shift tetap diperiksa terhadap mark.
    If Not FinalizeMedium(result, settings, doc) Then Exit Function
    If MinimumGroupValid(result, settings, doc) Then Set MinimumBlockSeed = result
End Function

Private Function MinimumGroupRefill(ByVal seed As AALayoutPlan, ByRef models() As AAModel, _
                                    ByRef indexes() As Long, ByVal count As Long, ByVal difference As Long, _
                                    ByVal totalLimit As Long, _
                                    ByVal settings As AASettings, ByVal doc As Document) As AALayoutPlan
    Dim result As AALayoutPlan, candidate As AALayoutPlan, placed As AAPlacement
    Dim allocation() As Long, tried As Object, forbidden As Collection, offsets As Object
    Dim i As Long, j As Long, index As Long, bestIndex As Long, least As Long, minimum As Long, maximum As Long
    Dim probes As Long, additions As Long, round As Long
    Set result = seed
    Set forbidden = ForbiddenFor(settings, doc)
    If settings.ModelMode = AA_MODEL_KISS_A Then Set forbidden = NominalForbiddenFor(settings, doc)
    Set offsets = CreateObject("Scripting.Dictionary"): offsets.Add "0", Array(0#, 0#)
    Do While additions < AA_GROUP_REFILL_ADDITIONS And probes < AA_GROUP_REFILL_PROBES And result.DesignCount < totalLimit
        ReDim allocation(0 To UBound(models))
        For Each placed In result.Placements
            allocation(placed.DesignIndex) = allocation(placed.DesignIndex) + 1
        Next placed
        Set tried = CreateObject("Scripting.Dictionary")
        Set candidate = Nothing
        For round = 1 To count
            bestIndex = -1: least = AA_MAX_OBJECTS_PER_SOURCE + 1
            For i = 0 To count - 1
                index = indexes(i)
                If Not tried.Exists(CStr(index)) And allocation(index) < least Then
                    bestIndex = index: least = allocation(index)
                End If
            Next i
            If bestIndex < 0 Then Exit For
            tried.Add CStr(bestIndex), True
            minimum = AA_MAX_OBJECTS_PER_SOURCE: maximum = 0
            For j = 0 To count - 1
                index = indexes(j): i = allocation(index)
                If index = bestIndex Then i = i + 1
                If i < minimum Then minimum = i
                If i > maximum Then maximum = i
            Next j
            If maximum - minimum <= difference And least < AA_MAX_OBJECTS_PER_SOURCE Then
                Set candidate = MinimumGroupBonus(result, models, bestIndex, forbidden, offsets, settings, doc, probes)
                If Not candidate Is Nothing Then Exit For
            End If
            If probes >= AA_GROUP_REFILL_PROBES Then Exit For
        Next round
        If candidate Is Nothing Then Exit Do
        Set result = candidate: additions = additions + 1
        AADebugWrite "Minimum block refill", "design=" & CStr(bestIndex) & "; allocated=" & _
            CStr(least + 1) & "; total=" & CStr(result.DesignCount)
    Loop
    AADebugWrite "Minimum block refill", "added=" & CStr(additions) & "; probes=" & CStr(probes) & _
        "/" & CStr(AA_GROUP_REFILL_PROBES)
    Set MinimumGroupRefill = result
End Function

Private Function MinimumGroupBonus(ByVal seed As AALayoutPlan, ByRef models() As AAModel, _
                                   ByVal index As Long, ByVal forbidden As Collection, ByVal offsets As Object, _
                                   ByVal settings As AASettings, ByVal doc As Document, ByRef probes As Long) As AALayoutPlan
    Dim model As AAModel, placed As AAPlacement, bonus As AAPlacement, origin As AAPlacement
    Dim candidate As AALayoutPlan, slots As New Collection, localSlots As New Collection, unique As Object
    Dim sx As Double, sy As Double, zx As Double, zy As Double, h As Double, v As Double, x As Double, y As Double
    Dim site As Long, i As Long, slot As Variant, key As String
    Set model = models(index)
    model.GetMinimumSteps AA_MODE_STRAIGHT, sx, sy
    model.GetMinimumSteps AA_MODE_ZIGZAG, zx, zy
    model.GetBoundaryGaps AA_MODE_STRAIGHT, h, v
    Set unique = CreateObject("Scripting.Dictionary")
    For Each placed In seed.Placements
        If placed.DesignIndex = index Then Set origin = placed
        unique(CStr(Round(placed.Footprint.LeftX, 4)) & "|" & CStr(Round(placed.Footprint.TopY, 4))) = True
    Next placed
    If origin Is Nothing Then Exit Function
    For Each placed In seed.Placements
        AAProgressPulse
        For site = 0 To 7
            MinimumLocalSlot placed.Footprint, site, sx, sy, zy, x, y
            key = CStr(Round(x, 4)) & "|" & CStr(Round(y, 4))
            If Not unique.Exists(key) Then
                unique.Add key, True: slots.Add Array(x, y)
            End If
        Next site
    Next placed
    ' Duluan seluruh tambahan langsung; jangan habiskan budget gerak lokal
    ' pada slot awal yang tertutup sementara masih ada slot bebas di bawah.
    For i = 1 To slots.Count
        AAProgressPulse
        slot = slots(i)
        If probes >= AA_GROUP_REFILL_PROBES Then Exit Function
        Set bonus = model.MakePlacement(index, origin.CutIndex, 0, _
            CDbl(slot(0)) + model.DesignInsetX, CDbl(slot(1)) - model.DesignInsetY)
        If Not origin.CreatesCutLine Then bonus.SuppressCutLine
        probes = probes + 1
        If CandidateFits(bonus, seed.Placements, forbidden, settings, offsets, h, v) Then
            Set candidate = MinimumGroupAdded(seed, Nothing, Nothing, bonus, settings, doc)
            If Not candidate Is Nothing Then Set MinimumGroupBonus = candidate: Exit Function
        ElseIf localSlots.Count < AA_GROUP_LOCAL_SLOTS Then
            If bonus.Footprint.LeftX >= 0# And bonus.Footprint.BottomY >= 0# And _
               bonus.Footprint.RightX <= settings.AreaWidthMM And bonus.Footprint.TopY <= settings.AreaHeightMM Then _
                localSlots.Add bonus
        End If
    Next i
    For Each bonus In localSlots
        Set candidate = MinimumGroupRelocate(seed, bonus, forbidden, offsets, settings, doc, probes)
        If Not candidate Is Nothing Then Set MinimumGroupBonus = candidate: Exit Function
        If probes >= AA_GROUP_REFILL_PROBES Then Exit Function
    Next bonus
End Function

Private Function MinimumGroupRelocate(ByVal seed As AALayoutPlan, ByVal bonus As AAPlacement, _
                                      ByVal forbidden As Collection, ByVal offsets As Object, _
                                      ByVal settings As AASettings, ByVal doc As Document, _
                                      ByRef probes As Long) As AALayoutPlan
    Dim original As AAPlacement, placed As AAPlacement, moved As AAPlacement
    Dim occupied As Collection, candidate As AALayoutPlan, localStep As Long, direction As Long
    Dim dx As Double, dy As Double, rx As Double, ry As Double
    For Each original In seed.Placements
        AAProgressPulse
        rx = (original.Footprint.RightX - original.Footprint.LeftX) * AA_GROUP_LOCAL_RADIUS
        ry = (original.Footprint.TopY - original.Footprint.BottomY) * AA_GROUP_LOCAL_RADIUS
        If bonus.Footprint.LeftX < original.Footprint.RightX + rx And _
           bonus.Footprint.RightX > original.Footprint.LeftX - rx And _
           bonus.Footprint.BottomY < original.Footprint.TopY + ry And _
           bonus.Footprint.TopY > original.Footprint.BottomY - ry Then
            Set occupied = New Collection
            For Each placed In seed.Placements
                If Not placed Is original Then occupied.Add placed
            Next placed
            If probes >= AA_GROUP_REFILL_PROBES Then Exit Function
            probes = probes + 1
            If CandidateFits(bonus, occupied, forbidden, settings, offsets, bonus.GapH, bonus.GapV) Then
                occupied.Add bonus
                For localStep = 1 To AA_GROUP_LOCAL_STEPS
                    For direction = 0 To 7
                        If probes >= AA_GROUP_REFILL_PROBES Then Exit Function
                        dx = 0#: dy = 0#
                        Select Case direction
                            Case 0, 4, 6: dx = rx * localStep / AA_GROUP_LOCAL_STEPS
                            Case 1, 5, 7: dx = -rx * localStep / AA_GROUP_LOCAL_STEPS
                        End Select
                        Select Case direction
                            Case 2, 4, 5: dy = ry * localStep / AA_GROUP_LOCAL_STEPS
                            Case 3, 6, 7: dy = -ry * localStep / AA_GROUP_LOCAL_STEPS
                        End Select
                        Set moved = original.CopyForPlan(): moved.Translate dx, dy
                        probes = probes + 1: AAProgressPulse
                        If CandidateFits(moved, occupied, forbidden, settings, offsets, moved.GapH, moved.GapV) Then
                            Set candidate = MinimumGroupAdded(seed, original, moved, bonus, settings, doc)
                            If Not candidate Is Nothing Then
                                AADebugWrite "Minimum block move", "design=" & CStr(original.DesignIndex) & _
                                    "; dx/dy=" & Format$(dx, "0.000000") & "/" & Format$(dy, "0.000000")
                                Set MinimumGroupRelocate = candidate: Exit Function
                            End If
                        End If
                    Next direction
                Next localStep
            End If
        End If
    Next original
End Function

Private Function MinimumGroupAdded(ByVal seed As AALayoutPlan, ByVal original As AAPlacement, _
                                   ByVal moved As AAPlacement, ByVal bonus As AAPlacement, _
                                   ByVal settings As AASettings, ByVal doc As Document) As AALayoutPlan
    Dim result As New AALayoutPlan, placed As AAPlacement
    result.Configure seed.LayoutMode: result.SetRegularPattern False
    For Each placed In seed.Placements
        AAProgressPulse
        If placed Is original Then result.Add moved.CopyForPlan() Else result.Add placed.CopyForPlan()
    Next placed
    result.Add bonus
    If MinimumGroupValid(result, settings, doc) Then
        AADebugWrite "Minimum block bonus", "design=" & CStr(bonus.DesignIndex) & _
            "; x/y=" & Format$(bonus.Footprint.LeftX, "0.000000") & "/" & Format$(bonus.Footprint.TopY, "0.000000")
        Set MinimumGroupAdded = result
    End If
End Function

Private Function MinimumGroupValid(ByVal plan As AALayoutPlan, ByVal settings As AASettings, _
                                    ByVal doc As Document) As Boolean
    Dim occupied As New Collection, forbidden As Collection, offsets As Object, placed As AAPlacement
    Set forbidden = ForbiddenFor(settings, doc)
    Set offsets = CreateObject("Scripting.Dictionary"): offsets.Add "0", Array(0#, 0#)
    For Each placed In plan.Placements
        AAProgressPulse
        If Not CandidateFits(placed, occupied, forbidden, settings, offsets, placed.GapH, placed.GapV) Then Exit Function
        occupied.Add placed
    Next placed
    MinimumGroupValid = FinalPlanFits(plan, settings, doc)
End Function

Private Function MinimumGroupAlignment(ByVal plan As AALayoutPlan) As Double
    Dim bins As Object, placed As AAPlacement, key As String, axis As Long, value As Double
    Set bins = CreateObject("Scripting.Dictionary")
    For Each placed In plan.Placements
        AAProgressPulse
        For axis = 0 To 1
            If axis = 0 Then value = placed.Footprint.LeftX Else value = placed.Footprint.TopY
            key = CStr(placed.DesignIndex) & ":" & CStr(axis) & ":" & CStr(Round(value, 2))
            If bins.Exists(key) Then
                MinimumGroupAlignment = MinimumGroupAlignment + CLng(bins(key))
                bins(key) = CLng(bins(key)) + 1
            Else
                bins.Add key, 1
            End If
        Next axis
    Next placed
End Function

Friend Function ForbiddenFor(ByVal settings As AASettings, _
                             ByVal sourceDocument As Document) As Collection
    Set ForbiddenFor = BuildForbidden(settings, sourceDocument)
End Function

Friend Function NominalForbiddenFor(ByVal settings As AASettings, _
                                    ByVal sourceDocument As Document) As Collection
    Set NominalForbiddenFor = BuildForbidden(settings, sourceDocument, 0#)
End Function

Friend Function FinalPlanFits(ByVal plan As AALayoutPlan, _
                              ByVal settings As AASettings, _
                              ByVal sourceDocument As Document) As Boolean
    Dim offsets As Object
    Dim areaIndex As Long, placement As AAPlacement, nominal As Collection

    Set offsets = CreateObject("Scripting.Dictionary")
    For areaIndex = 0 To plan.AreaCount - 1
        offsets.Add CStr(areaIndex), Array(0#, 0#)
    Next areaIndex
    FinalPlanFits = CenteredPlanFits(plan, offsets, _
        BuildForbidden(settings, sourceDocument), settings)
    If Not FinalPlanFits Then Exit Function
    For Each placement In plan.placements
        If placement.UsesMarkAreaAllowance Then
            If nominal Is Nothing Then Set nominal = NominalForbiddenFor(settings, sourceDocument)
            If Not PlacementMarksFit(placement, offsets, nominal) Then FinalPlanFits = False: Exit Function
        End If
    Next placement
End Function

Friend Function FinalizeMedium(ByVal plan As AALayoutPlan, ByVal settings As AASettings, _
                               ByVal sourceDocument As Document) As Boolean
    Dim offsets As Object, forbidden As Collection, nominal As Collection
    Dim placement As AAPlacement, shift As Variant
    Dim safe As Boolean
    If plan.designCount = 0 Then
        FinalizeMedium = True
        Exit Function
    End If
    Set forbidden = BuildForbidden(settings, sourceDocument)
    If settings.modelMode = AA_MODEL_KISS_A Then Set nominal = NominalForbiddenFor(settings, sourceDocument)
    Set offsets = CenterOffsets(plan, settings)
    safe = FindSafeOffsets(plan, offsets, forbidden, settings, nominal, True)
    If safe Then safe = CenteredPlanFits(plan, offsets, forbidden, settings)
    For Each placement In plan.placements
        shift = offsets(CStr(placement.areaIndex))
        placement.Translate CDbl(shift(0)), CDbl(shift(1))
    Next placement
    ' Saat pergeseran bersama tidak cukup, caller memperbaiki kandidat individual
    ' dari posisi yang sudah dipusatkan, tanpa mengubah Shape dokumen.
    FinalizeMedium = safe
End Function

Friend Function BetterPlan(ByVal candidate As AALayoutPlan, _
                           ByVal reference As AALayoutPlan) As Boolean
    Dim candidateBoundsArea As Double, referenceBoundsArea As Double
    Dim candidateBounds() As Double, referenceBounds() As Double
    Dim equivalentBounds As Boolean, areaIndex As Long
    Dim candidateRotation As Double, referenceRotation As Double
    If candidate.AreaCount <> reference.AreaCount Then
        BetterPlan = (candidate.AreaCount < reference.AreaCount)
    ElseIf candidate.designCount <> reference.designCount Then
        BetterPlan = (candidate.designCount > reference.designCount)
    ElseIf candidate.layoutMode = AA_MODE_MEDIUM And reference.layoutMode = AA_MODE_MEDIUM And _
           candidate.HasRegularPattern <> reference.HasRegularPattern Then
        BetterPlan = candidate.HasRegularPattern
    Else
        candidateBoundsArea = OccupiedBoundsArea(candidate, candidateBounds)
        referenceBoundsArea = OccupiedBoundsArea(reference, referenceBounds)
        ' Toleransi ini hanya untuk ranking Medium, bukan kelonggaran gap/bounds.
        equivalentBounds = (candidate.layoutMode = AA_MODE_MEDIUM And _
                            reference.layoutMode = AA_MODE_MEDIUM)
        If equivalentBounds Then
            For areaIndex = 0 To candidate.AreaCount - 1
                If Abs(candidateBounds(areaIndex, 0) - referenceBounds(areaIndex, 0)) > AA_PLAN_BOUNDS_TOLERANCE_MM Or _
                   Abs(candidateBounds(areaIndex, 1) - referenceBounds(areaIndex, 1)) > AA_PLAN_BOUNDS_TOLERANCE_MM Then
                    equivalentBounds = False
                    Exit For
                End If
            Next areaIndex
        End If
        If Not equivalentBounds And candidateBoundsArea <> referenceBoundsArea Then
            BetterPlan = (candidateBoundsArea < referenceBoundsArea)
        ElseIf equivalentBounds Then
            candidateRotation = TotalRotation(candidate)
            referenceRotation = TotalRotation(reference)
            If candidateRotation <> referenceRotation Then
                BetterPlan = (candidateRotation < referenceRotation)
            Else
                BetterPlan = HasFullerEarlierArea(candidate, reference)
            End If
        Else
            BetterPlan = HasFullerEarlierArea(candidate, reference)
        End If
    End If
End Function

Private Function TotalRotation(ByVal plan As AALayoutPlan) As Double
    Dim placement As AAPlacement, angle As Double
    For Each placement In plan.placements
        ' RotationDelta relatif terhadap sumber; 270 derajat setara putaran -90.
        angle = placement.RotationDelta
        angle = angle - 360# * Int(angle / 360#)
        If angle > 180# Then angle = 360# - angle
        TotalRotation = TotalRotation + angle
    Next placement
End Function

Private Function OccupiedBoundsArea(ByVal plan As AALayoutPlan, _
                                    ByRef bounds() As Double) As Double
    Dim lefts() As Double, bottoms() As Double
    Dim rights() As Double, tops() As Double
    Dim seen() As Boolean
    Dim placement As AAPlacement, box As AAGeometry
    Dim areaIndex As Long

    If plan.AreaCount = 0 Then Exit Function
    ReDim bounds(0 To plan.AreaCount - 1, 0 To 1)
    ReDim lefts(0 To plan.AreaCount - 1)
    ReDim bottoms(0 To plan.AreaCount - 1)
    ReDim rights(0 To plan.AreaCount - 1)
    ReDim tops(0 To plan.AreaCount - 1)
    ReDim seen(0 To plan.AreaCount - 1)
    For Each placement In plan.placements
        areaIndex = placement.areaIndex
        Set box = placement.collisionFootprint
        If seen(areaIndex) Then
            If box.leftX < lefts(areaIndex) Then lefts(areaIndex) = box.leftX
            If box.bottomY < bottoms(areaIndex) Then bottoms(areaIndex) = box.bottomY
            If box.rightX > rights(areaIndex) Then rights(areaIndex) = box.rightX
            If box.topY > tops(areaIndex) Then tops(areaIndex) = box.topY
        Else
            lefts(areaIndex) = box.leftX: bottoms(areaIndex) = box.bottomY
            rights(areaIndex) = box.rightX: tops(areaIndex) = box.topY
            seen(areaIndex) = True
        End If
    Next placement
    For areaIndex = 0 To plan.AreaCount - 1
        If seen(areaIndex) Then
            bounds(areaIndex, 0) = rights(areaIndex) - lefts(areaIndex)
            bounds(areaIndex, 1) = tops(areaIndex) - bottoms(areaIndex)
            OccupiedBoundsArea = OccupiedBoundsArea + _
                bounds(areaIndex, 0) * bounds(areaIndex, 1)
        End If
    Next areaIndex
End Function

Public Function BuildMinimum(ByRef models() As AAModel, _
                             ByRef designIndexes() As Long, ByVal designCount As Long, _
                             ByRef cutIndexes() As Long, ByVal cutCount As Long, _
                             ByRef quantities() As Long, ByVal settings As AASettings, _
                             ByVal sourceDocument As Document, _
                             Optional ByVal singleAreaOnly As Boolean = False, _
                             Optional ByVal reuseMarkChecks As Boolean = False) As AALayoutPlan
    Dim straightPlan As AALayoutPlan
    Dim zigZagPlan As AALayoutPlan
    Dim straightError As String
    Dim zigZagError As String

    AAProgressCandidate
    If AllActiveRectangles(models, designIndexes, designCount, quantities) Then
        AADebugCandidateStart "Minimum Straight"
        Set straightPlan = BuildOne(models, designIndexes, designCount, cutIndexes, _
            cutCount, quantities, settings, sourceDocument, AA_MODE_STRAIGHT, singleAreaOnly, reuseMarkChecks)
        AADebugCandidateEnd
        AADebugPlan "Minimum Straight", straightPlan
        Set BuildMinimum = straightPlan
        Exit Function
    End If

    On Error Resume Next
    AADebugCandidateStart "Minimum Straight"
    Set straightPlan = BuildOne(models, designIndexes, designCount, cutIndexes, _
        cutCount, quantities, settings, sourceDocument, AA_MODE_STRAIGHT, singleAreaOnly, reuseMarkChecks)
    If Err.number <> 0 Then straightError = Err.description
    Err.Clear
    On Error GoTo 0
    AAProgressCheckCancelled
    AADebugCandidateEnd
    AADebugPlan "Minimum Straight", straightPlan
    AAProgressCandidate
    AADebugCandidateStart "Minimum ZigZag"
    On Error Resume Next
    Set zigZagPlan = BuildOne(models, designIndexes, designCount, cutIndexes, _
        cutCount, quantities, settings, sourceDocument, AA_MODE_ZIGZAG, singleAreaOnly, reuseMarkChecks)
    If Err.number <> 0 Then zigZagError = Err.description
    Err.Clear
    On Error GoTo 0
    AAProgressCheckCancelled
    AADebugCandidateEnd
    AADebugPlan "Minimum ZigZag", zigZagPlan

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
        zigZagPlan.designCount > straightPlan.designCount Then
        Set BuildMinimum = zigZagPlan
    ElseIf zigZagPlan.AreaCount = straightPlan.AreaCount And _
        zigZagPlan.designCount = straightPlan.designCount Then
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
                          ByVal layoutMode As Long, ByVal singleAreaOnly As Boolean, _
                          ByVal reuseMarkChecks As Boolean) As AALayoutPlan
    Dim result As AALayoutPlan
    Dim forbidden As Collection, nominal As Collection
    Dim offsets As Object
    Dim nextOffsets As Object
    Dim placement As AAPlacement
    Dim shift As Variant
    Dim passIndex As Long
    Dim areaKey As Variant
    Dim perfStarted As Double, stageLabel As String
    Dim safe As Boolean, centeredSafe As Boolean

    Set forbidden = BuildForbidden(settings, sourceDocument)
    If settings.modelMode = AA_MODEL_KISS_A Then Set nominal = NominalForbiddenFor(settings, sourceDocument)
    ' Susun tanpa mark dahulu, lalu pusatkan seluruh isi tiap area.
    ' Coba geser seluruh isi area sebelum merencanakan ulang Quantity.
    For passIndex = 1 To AA_MAX_CENTERING_PASSES
        stageLabel = " layout=" & CStr(layoutMode) & "; pass=" & CStr(passIndex)
        perfStarted = AADebugPerfBegin()
        Set result = BuildOneRaw(models, designIndexes, designCount, cutIndexes, _
            cutCount, quantities, settings, layoutMode, forbidden, offsets, singleAreaOnly)
        AADebugStageEnd "Minimum raw build" & stageLabel, perfStarted
        If result.designCount = 0 Then
            Set BuildOne = result
            Exit Function
        End If
        Set nextOffsets = CenterOffsets(result, settings)
        If AADebugDetailsEnabled Then
            AADebugPlan "Minimum raw pass=" & CStr(passIndex), result
            For Each areaKey In nextOffsets.keys
                shift = nextOffsets(areaKey)
                AADebugWrite "Minimum", "center pass=" & CStr(passIndex) & "; area=" & CStr(areaKey) & _
                    "; shift X/Y=" & Format$(CDbl(shift(0)), "0.000000000") & "/" & _
                    Format$(CDbl(shift(1)), "0.000000000")
            Next areaKey
        End If
        perfStarted = AADebugPerfBegin()
        safe = FindSafeOffsets(result, nextOffsets, forbidden, settings, nominal, reuseMarkChecks)
        AADebugStageEnd "Minimum center search" & stageLabel, perfStarted
        If safe Then
            perfStarted = AADebugPerfBegin()
            centeredSafe = CenteredPlanFits(result, nextOffsets, forbidden, settings)
            AADebugStageEnd "Minimum center validation" & stageLabel, perfStarted
            If centeredSafe Then
                For Each placement In result.placements
                    shift = nextOffsets(CStr(placement.areaIndex))
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
                             ByVal offsets As Object, ByVal singleAreaOnly As Boolean) As AALayoutPlan
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
            If singleAreaOnly Then GoTo NextDesign
            If added = 0 And occupied.count > 0 Then
                areaIndex = areaIndex + 1
                Set occupied = New Collection
                added = FillArea(result, models(designIndex), designIndex, cutIndex, _
                    areaIndex, occupied, forbidden, settings, layoutMode, _
                    AA_MAX_OBJECTS_PER_SOURCE, offsets)
            End If
            If added = 0 Then Err.Raise vbObjectError + 5231, "AAPlanner", _
                "Obj. ID sumber tidak dapat muat dalam satu area."
            If cutCount = 1 And result.designCount > AA_MAX_OBJECTS_PER_SOURCE Then _
                Err.Raise vbObjectError + 5210, "AAPlanner", _
                "Satu Cut Line bersama melebihi batas 5000 objek."
        Else
            sourceCount = 0
            Do While remaining > 0
                added = FillArea(result, models(designIndex), designIndex, cutIndex, _
                    areaIndex, occupied, forbidden, settings, layoutMode, remaining, offsets)
                If added < remaining And occupied.count > 0 Then
                    added = added + FillMinimumNeighbors(result, models(designIndex), _
                        designIndex, cutIndex, areaIndex, occupied, forbidden, settings, _
                        layoutMode, remaining - added, offsets)
                End If
                If added = 0 And occupied.count = 0 Then Err.Raise _
                    vbObjectError + 5231, "AAPlanner", _
                    "Obj. ID sumber tidak dapat muat dalam satu area."
                remaining = remaining - added
                If singleAreaOnly And remaining > 0 Then Exit For
                sourceCount = sourceCount + added
                If sourceCount > AA_MAX_OBJECTS_PER_SOURCE Then Err.Raise _
                    vbObjectError + 5210, "AAPlanner", _
                    "Satu Design melebihi batas 5000 objek."
                If remaining > 0 Then
                    areaIndex = areaIndex + 1
                    Set occupied = New Collection
                End If
            Loop
            If cutCount = 1 And result.designCount > AA_MAX_OBJECTS_PER_SOURCE Then _
                Err.Raise vbObjectError + 5210, "AAPlanner", _
                "Satu Cut Line bersama melebihi batas 5000 objek."
        End If
NextDesign:
    Next roleIndex
    If requested And result.designCount = 0 Then Err.Raise _
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
    Dim markSlots As New Collection

    ' Semua rectangle selalu memakai Straight, termasuk sudut lengkung dan Custom.
    ' layoutMode ByVal: sumber lain tetap mengikuti pola rencana yang dibandingkan.
    If model.IsRectangle Then layoutMode = AA_MODE_STRAIGHT
    model.GetMinimumSteps layoutMode, horizontalStep, verticalStep
    model.GetBoundaryGaps layoutMode, horizontalGap, verticalGap
    BuildStartingRows occupied, designIndex, verticalGap, _
        cutIndex, model.collisionInsetY, settings.AreaHeightMM, startingRows
    gridTop = startingRows(0)
    yOffset = settings.AreaHeightMM - gridTop
    Do While yOffset + model.FootprintHeight <= settings.AreaHeightMM + AA_EDGE_TOLERANCE_MM
        rowAdded = False
        xOffset = 0#
        If layoutMode = AA_MODE_ZIGZAG And rowIndex Mod 2 = 1 Then _
            xOffset = horizontalStep / 2#
        Do While xOffset + model.FootprintWidth <= _
                 settings.AreaWidthMM + AA_EDGE_TOLERANCE_MM
            AAProgressPulse
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
            ElseIf markSlots.count < AA_MAX_MINIMUM_MARK_SLOTS And candidate.UsesContourGap Then
                If Not offsets Is Nothing Then
                    If offsets.Exists(CStr(areaIndex)) Then
                        If Not PlacementMarksFit(candidate, offsets, forbidden) Then markSlots.Add candidate
                    End If
                End If
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
    If FillArea < limit And markSlots.count > 0 Then
        FillArea = FillArea + FillMinimumMarkSlots(result, markSlots, occupied, forbidden, _
            settings, offsets, horizontalGap, verticalGap, limit - FillArea)
    End If
End Function

Private Function FillMinimumMarkSlots(ByVal result As AALayoutPlan, ByVal slots As Collection, _
                                      ByVal occupied As Collection, ByVal forbidden As Collection, _
                                      ByVal settings As AASettings, ByVal offsets As Object, _
                                      ByVal horizontalGap As Double, ByVal verticalGap As Double, _
                                      ByVal limit As Long) As Long
    Dim original As AAPlacement, candidate As AAPlacement, placed As AAPlacement
    Dim envelope As AAGeometry
    Dim probes As Long
    Dim perfStarted As Double
    If occupied.count = 0 Or limit <= 0 Then Exit Function
    If horizontalGap < 0# Or verticalGap < 0# Then Exit Function
    perfStarted = AADebugPerfBegin()
    ' Semua tetangga grid sudah tersedia. Pertahankan bounds agar pemulihan
    ' tidak mengubah offset pemusatan lalu mengulang masalah pada pass berikutnya.
    Set envelope = New AAGeometry
    For Each placed In occupied
        envelope.AddBounds placed.footprint.leftX, placed.footprint.bottomY, _
            placed.footprint.rightX, placed.footprint.topY
    Next placed
    For Each original In slots
        If AADebugDetailsEnabled Then AADebugMinimumSlotGap original, occupied
        Set candidate = ShiftMinimumMarkSlot(original, occupied, forbidden, settings, _
            offsets, horizontalGap, verticalGap, envelope, probes)
        If Not candidate Is Nothing Then
            result.Add candidate
            occupied.Add candidate
            FillMinimumMarkSlots = FillMinimumMarkSlots + 1
            If FillMinimumMarkSlots >= limit Then Exit For
        End If
    Next original
    If AADebugEnabled Then AADebugWrite "Minimum", "mark recovery: slots=" & CStr(slots.count) & _
        "; added=" & CStr(FillMinimumMarkSlots) & "; probes=" & CStr(probes) & "; max shift=2 mm"
    AADebugStageEnd "Minimum mark recovery", perfStarted
End Function

Private Function ShiftMinimumMarkSlot(ByVal original As AAPlacement, ByVal occupied As Collection, _
                                      ByVal forbidden As Collection, ByVal settings As AASettings, _
                                      ByVal offsets As Object, ByVal horizontalGap As Double, _
                                      ByVal verticalGap As Double, ByVal envelope As AAGeometry, _
                                      ByRef probes As Long) As AAPlacement
    Dim radii As Variant, radius As Variant, direction As Long
    Dim dx As Double, dy As Double, diagonal As Double
    Dim candidate As AAPlacement, box As AAGeometry
    Dim xs() As Double, ys() As Double, count As Long, i As Long
    Dim ring As Long, x As Long, y As Long, tested As Long, outside As Long
    Dim startProbes As Long, before As Variant, after As Variant, recovered As Boolean
    Set ShiftMinimumMarkSlot = Nothing
    recovered = False
    startProbes = probes
    If AADebugDetailsEnabled Then before = AADebugSlotSnapshot()
    ' Pertahankan pencarian delapan arah, lalu rapatkan kombinasi X/Y dekat
    ' slot semula. Batas salinan/pemeriksaan kurva tetap 256 per slot.
    ReDim xs(0 To 48 + 4 * AA_MINIMUM_MARK_FINE_STEPS * (AA_MINIMUM_MARK_FINE_STEPS + 1) - 1)
    ReDim ys(0 To UBound(xs))
    radii = Array(0.05, 0.1, 0.25, 0.5, 1#, 2#)
    For Each radius In radii
        diagonal = CDbl(radius) / Sqr(2#)
        For direction = 0 To 7
            dx = 0#: dy = 0#
            Select Case direction
                Case 0: dx = CDbl(radius)
                Case 1: dy = -CDbl(radius)
                Case 2: dx = -CDbl(radius)
                Case 3: dy = CDbl(radius)
                Case 4: dx = diagonal: dy = -diagonal
                Case 5: dx = -diagonal: dy = -diagonal
                Case 6: dx = -diagonal: dy = diagonal
                Case 7: dx = diagonal: dy = diagonal
            End Select
            xs(count) = dx: ys(count) = dy: count = count + 1
        Next direction
    Next radius
    For ring = 1 To AA_MINIMUM_MARK_FINE_STEPS
        For x = -ring To ring
            For y = -ring To ring
                If Abs(x) = ring Or Abs(y) = ring Then
                    xs(count) = x * AA_MINIMUM_MARK_FINE_STEP_MM
                    ys(count) = y * AA_MINIMUM_MARK_FINE_STEP_MM
                    count = count + 1
                End If
            Next y
        Next x
    Next ring
    Set box = original.footprint
    For i = 0 To count - 1
        dx = xs(i): dy = ys(i)
        AAProgressPulse
        probes = probes + 1
        If box.leftX + dx >= envelope.leftX - AA_EDGE_TOLERANCE_MM And _
           box.bottomY + dy >= envelope.bottomY - AA_EDGE_TOLERANCE_MM And _
           box.rightX + dx <= envelope.rightX + AA_EDGE_TOLERANCE_MM And _
           box.topY + dy <= envelope.topY + AA_EDGE_TOLERANCE_MM Then
            tested = tested + 1
            Set candidate = original.CopyForPlan
            candidate.Translate dx, dy
            ' True: posisi baru wajib menjaga gap juga terhadap Design yang sama.
            If CandidateFits(candidate, occupied, forbidden, settings, offsets, _
                horizontalGap, verticalGap, True) Then
                If AADebugDetailsEnabled Then AADebugWrite "Minimum", "mark recovered: design=" & _
                    CStr(original.designIndex) & "; area=" & CStr(original.areaIndex) & _
                    "; shift X/Y=" & Format$(dx, "0.000000") & "/" & Format$(dy, "0.000000")
                Set ShiftMinimumMarkSlot = candidate
                recovered = True
                Exit For
            End If
            If tested >= AA_MAX_MINIMUM_MARK_PROBES Then Exit For
        Else
            outside = outside + 1
        End If
    Next i
    If AADebugDetailsEnabled Then
        after = AADebugSlotSnapshot()
        AADebugWrite "Minimum", "mark slot: design=" & CStr(original.designIndex) & _
            "; area=" & CStr(original.areaIndex) & "; raw L/T=" & _
            Format$(box.leftX, "0.000000") & "/" & Format$(box.topY, "0.000000") & _
            "; recovered=" & CStr(recovered) & "; attempts=" & CStr(probes - startProbes) & _
            "; tested=" & CStr(tested) & "; envelope rejected=" & CStr(outside) & _
            "; bounds=" & CStr(CLng(after(0)) - CLng(before(0))) & _
            "; gap=" & CStr(CLng(after(1)) - CLng(before(1))) & _
            "; mark=" & CStr(CLng(after(2)) - CLng(before(2))) & _
            "; overlap=" & CStr(CLng(after(3)) - CLng(before(3))) & _
            "; limit reached=" & CStr(Not recovered And tested >= AA_MAX_MINIMUM_MARK_PROBES)
    End If
End Function

Private Function FillMinimumNeighbors(ByVal result As AALayoutPlan, ByVal model As AAModel, _
                                      ByVal designIndex As Long, ByVal cutIndex As Long, _
                                      ByVal areaIndex As Long, ByVal occupied As Collection, _
                                      ByVal forbidden As Collection, ByVal settings As AASettings, _
                                      ByVal layoutMode As Long, ByVal limit As Long, _
                                      ByVal offsets As Object) As Long
    Dim placed As AAPlacement, origin As AAPlacement, candidate As AAPlacement
    Dim box As AAGeometry, collision As AAGeometry
    Dim xs() As Double, ys() As Double, nx As Long, ny As Long
    Dim h As Double, v As Double, dx As Double, dy As Double
    Dim width As Double, height As Double, insetX As Double, insetY As Double
    Dim maxX As Double, minY As Double, x As Long, y As Long, pass As Long
    Dim probes As Long, before As Long, mixed As Boolean, key As String, tried As Object

    ' Tambahan terbatas untuk kuota finite campuran; fill bebas/grid satu sumber tetap.
    For Each placed In occupied
        If placed.designIndex <> designIndex Then mixed = True: Exit For
    Next placed
    If Not mixed Or limit <= 0 Then Exit Function
    If model.IsRectangle Then layoutMode = AA_MODE_STRAIGHT
    model.GetBoundaryGaps layoutMode, h, v
    If h < 0# Or v < 0# Then Exit Function
    Set origin = model.MakePlacement(designIndex, cutIndex, areaIndex, _
        model.DesignInsetX, settings.AreaHeightMM - model.DesignInsetY)
    If Not origin.UsesContourGap Then Exit Function
    Set collision = origin.collisionFootprint
    width = collision.width: height = collision.height
    insetX = collision.leftX - origin.footprint.leftX
    insetY = origin.footprint.topY - collision.topY
    maxX = settings.AreaWidthMM - model.FootprintWidth
    minY = model.FootprintHeight
    ReDim xs(0 To AA_MAX_MINIMUM_ANCHORS - 1)
    ReDim ys(0 To AA_MAX_MINIMUM_ANCHORS - 1)
    Set tried = CreateObject("Scripting.Dictionary")
    Do
        nx = 0: ny = 0: before = FillMinimumNeighbors
        AddMinimumAnchor xs, nx, 0#, 0#, maxX
        AddMinimumAnchor xs, nx, maxX, 0#, maxX
        AddMinimumAnchor ys, ny, settings.AreaHeightMM, minY, settings.AreaHeightMM
        AddMinimumAnchor ys, ny, minY, minY, settings.AreaHeightMM
        For Each placed In occupied
            Set box = placed.collisionFootprint
            AddMinimumAnchor xs, nx, box.leftX - insetX, 0#, maxX
            AddMinimumAnchor xs, nx, box.rightX - width - insetX, 0#, maxX
            AddMinimumAnchor ys, ny, box.topY + insetY, minY, settings.AreaHeightMM
            AddMinimumAnchor ys, ny, box.bottomY + height + insetY, minY, settings.AreaHeightMM
            ' Coba gap persis dahulu, lalu cadangan kecil untuk salinan kurva.
            For pass = 0 To 1
                dx = h + pass * AA_MINIMUM_NEIGHBOR_RESERVE_MM
                dy = v + pass * AA_MINIMUM_NEIGHBOR_RESERVE_MM
                If placed.gapH > h Then dx = placed.gapH + pass * AA_MINIMUM_NEIGHBOR_RESERVE_MM
                If placed.gapV > v Then dy = placed.gapV + pass * AA_MINIMUM_NEIGHBOR_RESERVE_MM
                AddMinimumAnchor xs, nx, box.rightX + dx - insetX, 0#, maxX
                AddMinimumAnchor xs, nx, box.leftX - dx - width - insetX, 0#, maxX
                AddMinimumAnchor ys, ny, box.bottomY - dy + insetY, minY, settings.AreaHeightMM
                AddMinimumAnchor ys, ny, box.topY + dy + height + insetY, minY, settings.AreaHeightMM
            Next pass
        Next placed
        ' Sumbu terurut: atas ke bawah, lalu kiri ke kanan. Anchor baru dibangun
        ' setelah placement diterima; posisi yang sudah dicoba tidak diulang.
        For y = ny - 1 To 0 Step -1
            For x = 0 To nx - 1
                key = CStr(xs(x)) & ":" & CStr(ys(y))
                If Not tried.Exists(key) Then
                    tried.Add key, True
                    probes = probes + 1
                    AAProgressPulse
                    Set candidate = model.MakePlacement(designIndex, cutIndex, areaIndex, _
                        xs(x) + model.DesignInsetX, ys(y) - model.DesignInsetY)
                    If CandidateFits(candidate, occupied, forbidden, settings, offsets, h, v, True) Then
                        result.Add candidate: occupied.Add candidate
                        FillMinimumNeighbors = FillMinimumNeighbors + 1
                        If FillMinimumNeighbors >= limit Then GoTo Finished
                    End If
                    If probes >= AA_MAX_MINIMUM_NEIGHBOR_PROBES Then GoTo Finished
                End If
            Next x
        Next y
    Loop While FillMinimumNeighbors > before
Finished:
    If AADebugEnabled Then AADebugWrite "Minimum", "neighbor fallback: design=" & _
        CStr(designIndex) & "; area=" & CStr(areaIndex) & "; added=" & _
        CStr(FillMinimumNeighbors) & "; probes=" & CStr(probes) & "/" & _
        CStr(AA_MAX_MINIMUM_NEIGHBOR_PROBES)
End Function

Private Sub AddMinimumAnchor(ByRef values() As Double, ByRef count As Long, _
                             ByVal value As Double, ByVal lower As Double, ByVal upper As Double)
    Dim i As Long, position As Long
    If value < lower - AA_EDGE_TOLERANCE_MM Or value > upper + AA_EDGE_TOLERANCE_MM Then Exit Sub
    If value < lower Then value = lower
    If value > upper Then value = upper
    position = count
    For i = 0 To count - 1
        If Abs(values(i) - value) <= AA_EDGE_TOLERANCE_MM Then Exit Sub
        If values(i) > value And position = count Then position = i
    Next i
    If count >= AA_MAX_MINIMUM_ANCHORS Then Exit Sub
    For i = count To position + 1 Step -1
        values(i) = values(i - 1)
    Next i
    values(position) = value: count = count + 1
End Sub

Private Sub BuildStartingRows(ByVal occupied As Collection, ByVal designIndex As Long, _
                              ByVal verticalGap As Double, ByVal cutIndex As Long, _
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
        ' Cut yang sama melanjutkan grid saat Cut itu pertama kali diisi.
        ' Hanya batas sumber SEBELUM template tersebut boleh memulai baris;
        ' batas anggota sendiri/sumber sesudahnya tidak mereset fase ZigZag.
        If cutIndex >= 0 And placed.cutIndex = cutIndex Then Exit For
        If placed.designIndex <> designIndex Then
            value = placed.collisionFootprint.bottomY - verticalGap + collisionInsetY
            If value >= 0# And value <= areaTop Then
                If Not unique.Exists(CStr(value)) Then unique.Add CStr(value), value
            End If
        End If
    Next placed
    ReDim values(0 To unique.count - 1)
    For Each key In unique.keys
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
                               ByVal horizontalGap As Double, ByVal verticalGap As Double, _
                               Optional ByVal traceGapFailure As Boolean = False) As Boolean
    Dim placed As AAPlacement
    Dim footprint As AAGeometry
    Dim collisionFootprint As AAGeometry
    Dim spacingFootprint As AAGeometry
    Dim placedSpacing As AAGeometry

    Set footprint = candidate.footprint
    Set collisionFootprint = candidate.collisionFootprint
    If footprint.leftX < -AA_EDGE_TOLERANCE_MM Or _
       footprint.bottomY < -AA_EDGE_TOLERANCE_MM Or _
       footprint.rightX > settings.AreaWidthMM + AA_EDGE_TOLERANCE_MM Or _
       footprint.topY > settings.AreaHeightMM + AA_EDGE_TOLERANCE_MM Then
        AADebugRejectSlot "bounds"
        Exit Function
    End If
    If Not offsets Is Nothing Then
        If offsets.Exists(CStr(candidate.areaIndex)) Then
            If Not PlacementMarksFit(candidate, offsets, forbidden, "Minimum slot") Then
                AADebugRejectSlot "mark"
                Exit Function
            End If
        End If
    End If
    For Each placed In occupied
        ' Acuan pasangan: Cut Line; standalone: siluet Design.
        Set spacingFootprint = collisionFootprint
        Set placedSpacing = placed.collisionFootprint
        If spacingFootprint.Overlaps(placedSpacing) Then
            AADebugRejectSlot "overlap"
            Exit Function
        End If
        ' Grid dan posisi tambahan memeriksa gap terhadap semua Design.
        If candidate.UsesContourGap And placed.UsesContourGap Then
            If Not candidate.HasClearance(placed) Then
                AADebugRejectSlot "gap"
                If traceGapFailure Then AADebugMinimumGapRejected candidate, placed, "Minimum local candidate"
                Exit Function
            End If
        Else
            If spacingFootprint.rightX + horizontalGap > placedSpacing.leftX + AA_EDGE_TOLERANCE_MM And _
               spacingFootprint.leftX - horizontalGap < placedSpacing.rightX - AA_EDGE_TOLERANCE_MM And _
               spacingFootprint.topY + verticalGap > placedSpacing.bottomY + AA_EDGE_TOLERANCE_MM And _
               spacingFootprint.bottomY - verticalGap < placedSpacing.topY - AA_EDGE_TOLERANCE_MM Then
                AADebugRejectSlot "gap"
                Exit Function
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
    For Each placement In plan.placements
        key = CStr(placement.areaIndex)
        Set footprint = placement.footprint
        If bounds.Exists(key) Then
            box = bounds(key)
            If footprint.leftX < box(0) Then box(0) = footprint.leftX
            If footprint.bottomY < box(1) Then box(1) = footprint.bottomY
            If footprint.rightX > box(2) Then box(2) = footprint.rightX
            If footprint.topY > box(3) Then box(3) = footprint.topY
            bounds(key) = box
        Else
            bounds.Add key, Array(footprint.leftX, footprint.bottomY, _
                footprint.rightX, footprint.topY)
        End If
    Next placement
    For Each key In bounds.keys
        box = bounds(key)
        result.Add key, Array((settings.AreaWidthMM - CDbl(box(0)) - _
            CDbl(box(2))) / 2#, (settings.AreaHeightMM - CDbl(box(1)) - _
            CDbl(box(3))) / 2#)
    Next key
    Set CenterOffsets = result
End Function

Private Function FindSafeOffsets(ByVal plan As AALayoutPlan, ByVal offsets As Object, _
                                 ByVal forbidden As Collection, _
                                 ByVal settings As AASettings, _
                                 ByVal nominal As Collection, _
                                 Optional ByVal reuseMarkChecks As Boolean = False) As Boolean
    Dim key As Variant
    Dim placement As AAPlacement
    Dim footprints As Collection
    Dim markFootprints As Collection
    Dim footprint As AAGeometry
    Dim shift As Variant
    Dim dx As Double, dy As Double
    Dim allSafe As Boolean, found As Boolean, extended As Boolean, fullMark As Boolean
    Dim markCache As Object

    allSafe = True
    For Each key In offsets.keys
        Set markCache = Nothing
        Set footprints = New Collection
        Set markFootprints = New Collection
        For Each placement In plan.placements
            If CStr(placement.areaIndex) = CStr(key) Then
                Set footprint = ShiftedFootprint(placement, offsets)
                footprints.Add footprint
                shift = offsets(key)
                Set footprint = placement.ForbiddenFootprint.placed( _
                    placement.ForbiddenFootprint.leftX + CDbl(shift(0)), _
                    placement.ForbiddenFootprint.topY + CDbl(shift(1)))
                markFootprints.Add footprint
            End If
        Next placement
        found = False: extended = False
        If Not nominal Is Nothing Then
            ' Hanya dua pencarian nominal memakai geometri dan mark yang sama.
            ' Cache lokal per area; tidak bertahan setelah translasi/fallback.
            If reuseMarkChecks Then Set markCache = NewShiftMarkCache()
            found = FindAreaShift(footprints, markFootprints, nominal, settings, dx, dy, 0#, markCache)
            If Not found Then
                found = FindAreaShift(footprints, markFootprints, nominal, settings, _
                    dx, dy, settings.MarkShiftAllowanceMM, markCache)
                extended = found
            End If
        End If
        Set markCache = Nothing
        fullMark = found And Not nominal Is Nothing
        If Not found Then found = FindAreaShift(footprints, markFootprints, forbidden, settings, dx, dy, 0#)
        If found Then
            shift = offsets(key)
            offsets(key) = Array(CDbl(shift(0)) + dx, CDbl(shift(1)) + dy)
        Else
            allSafe = False
        End If
        For Each placement In plan.placements
            If CStr(placement.areaIndex) = CStr(key) Then placement.SetMarkAreaAllowance extended
        Next placement
        If AADebugEnabled Then AADebugWrite "Mark shift", "area=" & CStr(key) & _
            "; found=" & CStr(found) & "; full mark=" & CStr(fullMark) & _
            "; allowance=" & CStr(extended) & "; extra X/Y=" & _
            Format$(dx, "0.000000") & "/" & Format$(dy, "0.000000")
    Next key
    FindSafeOffsets = allSafe
End Function

Private Function FindAreaShift(ByVal footprints As Collection, _
                               ByVal markFootprints As Collection, _
                               ByVal forbidden As Collection, ByVal settings As AASettings, _
                               ByRef bestX As Double, ByRef bestY As Double, _
                               ByVal areaAllowance As Double, _
                               Optional ByVal markCache As Object = Nothing) As Boolean
    Dim footprint As AAGeometry
    Dim leftX As Double, bottomY As Double, rightX As Double, topY As Double
    Dim xs() As Double, ys() As Double
    Dim i As Long, j As Long, probes As Long
    Dim distance As Double, bestDistance As Double
    Dim found As Boolean
    Dim perfStarted As Double, cacheHitsBefore As Long

    perfStarted = AADebugPerfBegin()
    cacheHitsBefore = ShiftMarkCacheHits(markCache)
    bestX = 0#
    bestY = 0#
    If AreaShiftFits(footprints, markFootprints, 0#, 0#, forbidden, settings, 0#, markCache) Then
        found = True
        GoTo SearchFinished
    End If
    Set footprint = footprints(1)
    leftX = footprint.leftX
    bottomY = footprint.bottomY
    rightX = footprint.rightX
    topY = footprint.topY
    For Each footprint In footprints
        If footprint.leftX < leftX Then leftX = footprint.leftX
        If footprint.bottomY < bottomY Then bottomY = footprint.bottomY
        If footprint.rightX > rightX Then rightX = footprint.rightX
        If footprint.topY > topY Then topY = footprint.topY
    Next footprint
    ' Ruang ini hanya untuk translasi: ukuran kelompok sendiri wajib muat
    ' di area nominal, sehingga allowance tidak membuka kapasitas baru.
    If rightX - leftX > settings.AreaWidthMM + 2# * AA_EDGE_TOLERANCE_MM Or _
       topY - bottomY > settings.AreaHeightMM + 2# * AA_EDGE_TOLERANCE_MM Then GoTo SearchFinished
    ' Kandidat 0,2 mm dan posisi tepat di batas ruang gerak; urut jarak dari pusat.
    BuildShiftValues -leftX - areaAllowance, settings.AreaWidthMM - rightX + areaAllowance, xs
    BuildShiftValues -bottomY - areaAllowance, settings.AreaHeightMM - topY + areaAllowance, ys
    bestDistance = settings.AreaWidthMM ^ 2 + settings.AreaHeightMM ^ 2 + 1#
    For i = LBound(xs) To UBound(xs)
        AAProgressPulse
        If xs(i) ^ 2 > bestDistance Then Exit For
        For j = LBound(ys) To UBound(ys)
            distance = xs(i) ^ 2 + ys(j) ^ 2
            If distance >= bestDistance Then Exit For
            probes = probes + 1
            If probes > AA_MAX_SHIFT_PROBES_PER_AREA Then GoTo SearchFinished
            If AreaShiftFits(footprints, markFootprints, xs(i), ys(j), forbidden, settings, areaAllowance, markCache) Then
                found = True
                bestDistance = distance
                bestX = xs(i)
                bestY = ys(j)
            End If
        Next j
    Next i
SearchFinished:
    FindAreaShift = found
    If AADebugEnabled Then AADebugWrite "Mark search", "allowance=" & CStr(areaAllowance) & _
        "; objects=" & CStr(footprints.count) & "; marks=" & CStr(forbidden.count) & _
        "; probes=" & CStr(probes) & "; found=" & CStr(found) & _
        "; cache=" & CStr(Not markCache Is Nothing) & _
        "; hits=" & CStr(ShiftMarkCacheHits(markCache) - cacheHitsBefore)
    AADebugStageEnd "Mark shift search", perfStarted
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
                               ByVal settings As AASettings, ByVal areaAllowance As Double, _
                               Optional ByVal markCache As Object = Nothing) As Boolean
    Dim footprint As AAGeometry

    For Each footprint In footprints
        If footprint.leftX + dx < -areaAllowance - AA_EDGE_TOLERANCE_MM Or _
           footprint.bottomY + dy < -areaAllowance - AA_EDGE_TOLERANCE_MM Or _
           footprint.rightX + dx > settings.AreaWidthMM + areaAllowance + AA_EDGE_TOLERANCE_MM Or _
           footprint.topY + dy > settings.AreaHeightMM + areaAllowance + AA_EDGE_TOLERANCE_MM Then Exit Function
    Next footprint
    ' Batas area selalu diperiksa ulang: allowance tidak menjadi bagian cache mark.
    AreaShiftFits = ShiftMarksFit(markFootprints, dx, dy, forbidden, markCache)
End Function

Private Function NewShiftMarkCache() As Object
    Dim result As Object
    On Error Resume Next
    Set result = CreateObject("Scripting.Dictionary")
    result.Add "hits", 0&
    If Err.number <> 0 Then Set result = Nothing
    Err.Clear
    Set NewShiftMarkCache = result
End Function

Private Function ShiftMarkCacheHits(ByVal cache As Object) As Long
    On Error Resume Next
    If Not cache Is Nothing Then ShiftMarkCacheHits = CLng(cache.item("hits"))
    Err.Clear
End Function

' Cache hanya boleh dipakai selama markFootprints dan forbidden tetap sama.
' String memilih entry; dua Double asli wajib sama persis sebelum reuse.
Private Function ShiftMarksFit(ByVal markFootprints As Collection, _
                               ByVal dx As Double, ByVal dy As Double, _
                               ByVal forbidden As Collection, ByVal cache As Object) As Boolean
    Dim footprint As AAGeometry, key As String, entry As Variant
    Dim cachedX As Double, cachedY As Double, cachedResult As Boolean
    Dim hit As Boolean, result As Boolean
    If Not cache Is Nothing Then
        key = CStr(dx) & "|" & CStr(dy)
        ' Kegagalan cache kembali ke predicate; error geometri tidak ditelan.
        On Error Resume Next
        If cache.Exists(key) Then entry = cache.item(key)
        If IsArray(entry) Then
            cachedX = CDbl(entry(0)): cachedY = CDbl(entry(1)): cachedResult = CBool(entry(2))
            If Err.number = 0 Then hit = (cachedX = dx And cachedY = dy)
        End If
        If hit Then cache.item("hits") = ShiftMarkCacheHits(cache) + 1
        Err.Clear
        On Error GoTo 0
        If hit Then
            ShiftMarksFit = cachedResult
            Exit Function
        End If
    End If
    result = True
    For Each footprint In markFootprints
        If Not MarksFit(footprint, dx, dy, forbidden) Then
            result = False
            Exit For
        End If
    Next footprint
    If Not cache Is Nothing Then
        On Error Resume Next
        ' Termasuk satu entry statistik; penuh tetap melayani hit yang ada.
        If cache.Exists(key) Or cache.count <= AA_MAX_SHIFT_CACHE_ENTRIES Then _
            cache.item(key) = Array(dx, dy, result)
        Err.Clear
        On Error GoTo 0
    End If
    ShiftMarksFit = result
End Function

Private Function ShiftedFootprint(ByVal placement As AAPlacement, _
                                  ByVal offsets As Object) As AAGeometry
    Dim shift As Variant
    Dim result As AAGeometry
    Dim footprint As AAGeometry
    shift = offsets(CStr(placement.areaIndex))
    Set footprint = placement.footprint
    Set result = New AAGeometry
    ' Pemusatan/batas area hanya memerlukan angka, bukan salinan kurva.
    result.AddBounds footprint.leftX + CDbl(shift(0)), _
        footprint.bottomY + CDbl(shift(1)), footprint.rightX + CDbl(shift(0)), _
        footprint.topY + CDbl(shift(1))
    Set ShiftedFootprint = result
End Function

Private Function CenteredPlanFits(ByVal plan As AALayoutPlan, ByVal offsets As Object, _
                                  ByVal forbidden As Collection, _
                                  ByVal settings As AASettings) As Boolean
    Dim placement As AAPlacement
    Dim footprint As AAGeometry
    Dim allowance As Double, bounds As Object, box As Variant, key As Variant
    Set bounds = CreateObject("Scripting.Dictionary")
    For Each placement In plan.placements
        Set footprint = ShiftedFootprint(placement, offsets)
        allowance = 0#
        If placement.UsesMarkAreaAllowance Then allowance = settings.MarkShiftAllowanceMM
        If footprint.leftX < -allowance - AA_EDGE_TOLERANCE_MM Or _
           footprint.bottomY < -allowance - AA_EDGE_TOLERANCE_MM Or _
           footprint.rightX > settings.AreaWidthMM + allowance + AA_EDGE_TOLERANCE_MM Or _
           footprint.topY > settings.AreaHeightMM + allowance + AA_EDGE_TOLERANCE_MM Then Exit Function
        If Not PlacementMarksFit(placement, offsets, forbidden) Then Exit Function
        key = CStr(placement.areaIndex)
        If bounds.Exists(key) Then
            box = bounds(key)
            If footprint.leftX < box(0) Then box(0) = footprint.leftX
            If footprint.bottomY < box(1) Then box(1) = footprint.bottomY
            If footprint.rightX > box(2) Then box(2) = footprint.rightX
            If footprint.topY > box(3) Then box(3) = footprint.topY
            bounds(key) = box
        Else
            bounds.Add key, Array(footprint.leftX, footprint.bottomY, footprint.rightX, footprint.topY)
        End If
    Next placement
    For Each key In bounds.keys
        box = bounds(key)
        If CDbl(box(2)) - CDbl(box(0)) > settings.AreaWidthMM + 2# * AA_EDGE_TOLERANCE_MM Or _
           CDbl(box(3)) - CDbl(box(1)) > settings.AreaHeightMM + 2# * AA_EDGE_TOLERANCE_MM Then Exit Function
    Next key
    CenteredPlanFits = True
End Function

Private Function PlacementMarksFit(ByVal placement As AAPlacement, _
                                   ByVal offsets As Object, _
                                    ByVal forbidden As Collection, _
                                    Optional ByVal traceLabel As String = vbNullString) As Boolean
    Dim shift As Variant
    shift = offsets(CStr(placement.areaIndex))
    If Len(traceLabel) > 0 And AADebugDetailsEnabled Then
        traceLabel = traceLabel & "; design=" & CStr(placement.designIndex) & _
            "; area=" & CStr(placement.areaIndex)
    Else
        traceLabel = vbNullString
    End If
    PlacementMarksFit = MarksFit(placement.ForbiddenFootprint, _
        CDbl(shift(0)), CDbl(shift(1)), forbidden, traceLabel)
End Function

Private Function MarksFit(ByVal footprint As AAGeometry, _
                          ByVal dx As Double, ByVal dy As Double, _
                           ByVal forbidden As Collection, _
                           Optional ByVal traceLabel As String = vbNullString) As Boolean
    Dim obstacle As AAGeometry
    Dim moved As AAGeometry
    Dim markIndex As Long
    For Each obstacle In forbidden
        markIndex = markIndex + 1
        ' Salin kurva hanya jika bounding box dekat dengan salah satu mark.
        If footprint.leftX + dx < obstacle.rightX - AA_EDGE_TOLERANCE_MM And _
           footprint.rightX + dx > obstacle.leftX + AA_EDGE_TOLERANCE_MM And _
           footprint.bottomY + dy < obstacle.topY - AA_EDGE_TOLERANCE_MM And _
           footprint.topY + dy > obstacle.bottomY + AA_EDGE_TOLERANCE_MM Then
            If moved Is Nothing Then Set moved = footprint.placed( _
                footprint.leftX + dx, footprint.topY + dy)
            If moved.Overlaps(obstacle) Then
                If Len(traceLabel) > 0 Then AADebugMarkRejected traceLabel, footprint, moved, obstacle, dx, dy, markIndex
                Exit Function
            End If
        End If
    Next obstacle
    MarksFit = True
End Function

Private Function BuildForbidden(ByVal settings As AASettings, _
                                ByVal sourceDocument As Document, _
                                Optional ByVal markTolerance As Double = AA_MARK_TOLERANCE_MM) As Collection
    Dim result As New Collection
    Dim markLayer As Layer
    Select Case settings.modelMode
        Case AA_MODEL_CUSTOM
            Set result = settings.CustomForbidden
        Case AA_MODEL_KISS_A
            Set markLayer = KissMarkLayer(sourceDocument)
            AddKissCorner result, 6#, 10#, settings, sourceDocument, markLayer, markTolerance
            AddKissCorner result, 6#, 475#, settings, sourceDocument, markLayer, markTolerance
            AddKissCorner result, 319#, 10#, settings, sourceDocument, markLayer, markTolerance
            AddKissCorner result, 319#, 475#, settings, sourceDocument, markLayer, markTolerance
            AddMark result, 162.5, 11.5, 12#, settings, sourceDocument, False, markTolerance
        Case AA_MODEL_DIE_A
            AddMark result, 11#, 11#, 11#, settings, sourceDocument, True, markTolerance
            AddMark result, 11#, 474#, 11#, settings, sourceDocument, True, markTolerance
            AddMark result, 314#, 11#, 11#, settings, sourceDocument, True, markTolerance
            AddMark result, 314#, 474#, 11#, settings, sourceDocument, True, markTolerance
    End Select
    Set BuildForbidden = result
End Function

Private Function KissMarkLayer(ByVal sourceDocument As Document) As Layer
    Dim candidate As Layer, found As Layer
    If sourceDocument Is Nothing Then Exit Function
    For Each candidate In sourceDocument.ActivePage.Layers
        AAProgressPulse
        If StrComp(candidate.name, "scpro2_regmarks", vbTextCompare) = 0 Then
            ' Lebih dari satu layer sama-sama ambigu: gunakan pusat bawaan.
            If Not found Is Nothing Then Exit Function
            Set found = candidate
        End If
    Next candidate
    Set KissMarkLayer = found
End Function

Private Sub AddKissCorner(ByVal result As Collection, ByVal nominalX As Double, ByVal nominalY As Double, _
                          ByVal settings As AASettings, ByVal sourceDocument As Document, _
                          ByVal markLayer As Layer, ByVal markTolerance As Double)
    Dim candidate As Shape, sourcePage As page, count As Long
    Dim originX As Double, originY As Double, x As Double, y As Double, halfSearch As Double
    Dim foundX As Double, foundY As Double, centerX As Double, centerY As Double
    x = nominalX: y = nominalY
    If Not markLayer Is Nothing Then
        Set sourcePage = sourceDocument.ActivePage
        originX = sourcePage.centerX - settings.PresetPageWidthMM / 2#
        originY = sourcePage.centerY - settings.PresetPageHeightMM / 2#
        halfSearch = AA_KISS_ANCHOR_SEARCH_MM / 2# + AA_KISS_ANCHOR_TOLERANCE_MM
        ' Seperti WhiteSpace AutoSensor: objek langsung saja, group dihitung
        ' satu kandidat, tanpa filter nama atau penelusuran anggota group.
        For Each candidate In markLayer.shapes
            AAProgressPulse
            centerX = candidate.centerX: centerY = candidate.centerY
            If Abs(centerX - originX - nominalX) <= halfSearch And _
               Abs(centerY - originY - nominalY) <= halfSearch Then
                count = count + 1
                foundX = centerX - originX: foundY = centerY - originY
            End If
        Next candidate
        If count = 1 Then x = foundX: y = foundY
    End If
    If AADebugDetailsEnabled Then AADebugWrite "Kiss corner", _
        "nominal X/Y=" & CStr(nominalX) & "/" & CStr(nominalY) & "; candidates=" & CStr(count) & _
        "; anchor=" & CStr(count = 1) & "; center X/Y=" & CStr(x) & "/" & CStr(y)
    If count = 1 Then
        AddMark result, x, y, AA_KISS_CORNER_DIAMETER_MM, settings, sourceDocument, True, markTolerance
    Else
        AddMark result, nominalX, nominalY, AA_KISS_CORNER_FALLBACK_MM, settings, sourceDocument, False, markTolerance
    End If
End Sub

Private Sub AddMark(ByVal result As Collection, ByVal centerX As Double, _
                    ByVal centerY As Double, ByVal sizeMM As Double, _
                    ByVal settings As AASettings, ByVal sourceDocument As Document, _
                    Optional ByVal elliptical As Boolean = False, _
                    Optional ByVal markTolerance As Double = AA_MARK_TOLERANCE_MM)
    Dim mark As AAGeometry
    Dim localX As Double
    Dim localY As Double
    localX = centerX - (settings.PresetPageWidthMM - settings.AreaWidthMM) / 2#
    localY = centerY - (settings.PresetPageHeightMM - settings.AreaHeightMM) / 2#
    Set mark = New AAGeometry
    ' Abaikan penetrasi hingga 0,2 mm dari tepi mark, bukan dari tepi object/area.
    If elliptical Then
        mark.AddEllipse localX, localY, sizeMM / 2# - markTolerance, _
            sizeMM / 2# - markTolerance, sourceDocument
    Else
        mark.AddRectangle localX - sizeMM / 2# + markTolerance, _
            localY - sizeMM / 2# + markTolerance, _
            localX + sizeMM / 2# - markTolerance, _
            localY + sizeMM / 2# - markTolerance, sourceDocument
    End If
    result.Add mark
End Sub
