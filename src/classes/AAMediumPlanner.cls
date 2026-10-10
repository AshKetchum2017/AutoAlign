Option Explicit

Private Const AA_SHIFT_FACTOR As Double = 0.2
Private Const AA_SHIFT_STEP_MM As Double = 1#
Private Const AA_EDGE_TOLERANCE_MM As Double = 0.0001
Private Const AA_MAX_OBJECTS_PER_SOURCE As Long = 5000
Private Const AA_MAX_PROBES_PER_AREA As Long = 50000
Private Const AA_MAX_SLOTS_PER_AREA As Long = 50000
Private Const AA_MAX_CONTACT_ANCHORS As Long = 16
Private Const AA_MAX_REFINEMENT_PROBES_PER_AREA As Long = 2048
Private Const AA_MAX_REFINEMENT_SLOTS As Long = 2048
Private Const AA_COMPACTION_TOLERANCE_MM As Double = 0.01
' Cadangan posisi Compact untuk pembulatan salinan/pemusatan kurva CorelDRAW.
Private Const AA_COMPACTION_RESERVE_MM As Double = 0.001
Private Const AA_MAX_COMPACTION_PROBES_PER_AREA As Long = 4096
Private Const AA_COMPACTION_PASSES As Long = 2
Private Const AA_MAX_ALIGNMENT_PROBES_PER_AREA As Long = 256
Private Const AA_ALIGNMENT_PASSES As Long = 2
Private Const AA_MAX_PAIR_CONTOUR_EDGES As Long = 32
' Lane terbatas untuk kurva sederhana (node sedikit, hasil flatten lebih banyak).
' Tidak membuka pencarian contact scan 64 langkah untuk semua kurva kompleks.
Private Const AA_MAX_CURVED_PAIR_EDGES As Long = 512
Private Const AA_MAX_CURVED_PAIR_NODES As Long = 12
Private Const AA_CURVED_PHASE_STEPS As Long = 4
Private Const AA_PAIR_CONTACT_SCAN_STEPS As Long = 64
Private Const AA_MAX_PHASE_SLOTS_PER_AREA As Long = 5000
Private Const AA_PAIR_INSET_RESERVE_MM As Double = 0.001

Private mOccupiedIndex As AAOccupiedIndex

' Diaktifkan hanya oleh satu pencarian kelompok /n. Tidak lintas Process/group.
' Cache menyimpan urutan FillArea SEBELUM centering/repair/Compact.
Private Const AA_MAX_CACHED_PLACEMENTS As Long = 8192
Private mQuotaFills As Object
Private mCachedPlacements As Long
Private mFillHits As Long, mFillMisses As Long, mFillReused As Long

' Keputusan gap identik selama satu kelompok quota; kontur rotasi immutable.
' Posisi absolut dibandingkan persis agar pembulatan translasi tetap diperiksa.
Private Const AA_QUOTA_GAP_BUCKETS As Long = 65521
Private Const AA_QUOTA_GAP_WAYS As Long = 4
Private Const AA_QUOTA_GAP_CACHE_SIZE As Long = AA_QUOTA_GAP_BUCKETS * AA_QUOTA_GAP_WAYS
Private Const AA_QUOTA_GAP_HASH_LIMIT As Double = 1000000#
Private mGapOrientations() As AARotationVariant, mGapOthers() As AAContour
' x, y, blocker left/top, effective gap H/V.
Private mGapKeys() As Double, mGapResults() As Boolean
Private mGapNextVictim() As Byte
Private mGapCacheReady As Boolean, mGapCacheDisabled As Boolean
Private mGapHits As Double, mGapMisses As Double, mGapBypasses As Double
Private mGapCollisions As Double, mGapStored As Long

Friend Sub BeginQuotaSearch()
    Set mQuotaFills = CreateObject("Scripting.Dictionary")
    mCachedPlacements = 0
    Erase mGapOrientations: Erase mGapOthers: Erase mGapKeys: Erase mGapResults
    Erase mGapNextVictim
    mGapCacheReady = False: mGapCacheDisabled = False: mGapStored = 0
    Set mOccupiedIndex = Nothing
End Sub

Public Function Build(ByRef models() As AAModel, ByRef designIndexes() As Long, _
                      ByVal designCount As Long, ByRef cutIndexes() As Long, _
                      ByVal cutCount As Long, ByRef quantities() As Long, _
                      ByVal settings As AASettings, ByVal sourceDocument As Document, _
                      ByVal sharedPlanner As AAPlanner, _
                      Optional ByVal singleAreaOnly As Boolean = False, _
                      Optional ByVal capacityTarget As Long = -1) As AALayoutPlan
    Dim best As AALayoutPlan, minimum As AALayoutPlan, candidate As AALayoutPlan
    Dim forbidden As Collection, angles As Collection, angle As Variant
    Dim firstVariant As AARotationVariant
    Dim firstIndex As Long, i As Long, pattern As Long, lastError As String
    Dim pairAngle As Long, rowMode As Long, patternMode As Long, pairCutIndex As Long
    Dim hasZigZag As Boolean, pairEligible As Boolean
    Dim candidateLabel As String, bestLabel As String, buildError As String
    Dim perfStarted As Double, capacityValidated As Boolean
    mFillHits = 0: mFillMisses = 0: mFillReused = 0
    mGapHits = 0#: mGapMisses = 0#: mGapBypasses = 0#: mGapCollisions = 0#
    ' Capacity hanya untuk quota finite pada satu area; caller biasa memakai -1.
    If Not singleAreaOnly Then capacityTarget = -1
    Set forbidden = sharedPlanner.ForbiddenFor(settings, sourceDocument)
    perfStarted = AADebugPerfBegin()
    On Error Resume Next
    Set minimum = sharedPlanner.BuildMinimum(models, designIndexes, designCount, _
        cutIndexes, cutCount, quantities, settings, sourceDocument, singleAreaOnly)
    If Err.Number <> 0 Then lastError = Err.Description
    Err.Clear
    On Error GoTo 0
    If Not minimum Is Nothing Then
        Set candidate = AdaptMinimum(minimum, models)
        If ValidPlan(candidate, forbidden, settings, "Baseline Minimum") Then
            ConsiderCandidate candidate, best, "Baseline Minimum", bestLabel, sharedPlanner
        Else
            AADebugWrite "Medium", "Baseline Minimum rejected: gap, collision, area or mark"
        End If
    Else
        AADebugWrite "Medium", "Baseline Minimum unavailable: " & lastError
    End If
    AADebugStageEnd "Baseline Minimum + validation", perfStarted
    If capacityTarget > 0 Then
        If Not best Is Nothing Then
            If best.MatchesQuotas(designIndexes, designCount, quantities, capacityTarget) Then
                capacityValidated = True
                GoTo BuildComplete
            End If
        End If
    End If
    If designCount = 1 Then
        firstIndex = designIndexes(0)
        If quantities(firstIndex) < 0 And Not models(firstIndex).IsRectangle Then
            pairEligible = models(firstIndex).IsTriangle
            If Not pairEligible Then
                On Error Resume Next
                Set firstVariant = models(firstIndex).RotationVariant(0#)
                If Err.Number = 0 And Not firstVariant Is Nothing Then
                    pairEligible = (firstVariant.Contour.EdgeCount <= AA_MAX_PAIR_CONTOUR_EDGES)
                End If
                Err.Clear
                On Error GoTo 0
            End If
            If AADebugEnabled Then AADebugWrite "Pair", "eligible=" & CStr(pairEligible) & _
                "; contour edge limit=" & CStr(AA_MAX_PAIR_CONTOUR_EDGES)
            If pairEligible Then
                pairCutIndex = -1
                If cutCount = 1 Then pairCutIndex = cutIndexes(0)
                For pairAngle = 0 To 315 Step 45
                    For rowMode = 0 To 1
                        For patternMode = 0 To 3
                            If patternMode = 1 And pairAngle <> 0 And pairAngle <> 90 Then GoTo NextPairPattern
                            If patternMode = 2 Then
                                If pairAngle <> 0 Or rowMode <> 0 Or _
                                   Not models(firstIndex).IsTriangle Then GoTo NextPairPattern
                            End If
                            ' Offset Y membuat urutan pasangan berpengaruh pada kurva asimetris.
                            ' Coba kedua urutan di semua sudut siku, termasuk sumber yang sudah diputar.
                            If patternMode = 3 Then
                                If rowMode <> 0 Or pairAngle Mod 90 <> 0 Then Exit For
                            End If
                            Set candidate = Nothing
                            candidateLabel = "Pair angle=" & CStr(pairAngle) & _
                                " row=" & CStr(rowMode) & " pattern=" & CStr(patternMode)
                            AAProgressCandidate
                            AADebugCandidateStart candidateLabel
                            buildError = vbNullString
                            On Error Resume Next
                            Set candidate = BuildOppositeLattice(models(firstIndex), firstIndex, _
                                pairCutIndex, settings, forbidden, CDbl(pairAngle), rowMode, patternMode)
                            If Err.Number <> 0 Then
                                buildError = "Error " & CStr(Err.Number) & ": " & Err.Description
                                lastError = Err.Description
                            End If
                            Err.Clear
                            On Error GoTo 0
                            AADebugCandidateEnd
                            If Len(buildError) > 0 Then AADebugWrite "Pair", candidateLabel & " -> " & buildError
                            Set candidate = FinalizeCandidate(candidate, models, forbidden, settings, _
                                sourceDocument, sharedPlanner, candidateLabel)
                            ConsiderCandidate candidate, best, candidateLabel, bestLabel, sharedPlanner
NextPairPattern:
                        Next patternMode
                    Next rowMode
                Next pairAngle
            ElseIf Not firstVariant Is Nothing Then
                If AADebugEnabled Then AADebugWrite "Pair", "curved lane eligible=" & _
                    CStr(CurvedPairEligible(firstVariant)) & "; nodes=" & _
                    CStr(firstVariant.Contour.Anchors.Count) & "; edges=" & CStr(firstVariant.Contour.EdgeCount)
                If CurvedPairEligible(firstVariant) Then
                    pairCutIndex = -1
                    If cutCount = 1 Then pairCutIndex = cutIndexes(0)
                    ConsiderCurvedPairs models(firstIndex), firstIndex, pairCutIndex, _
                        settings, forbidden, models, sourceDocument, sharedPlanner, best, bestLabel
                End If
            End If
        End If
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
            candidateLabel = "Pass angle=" & CStr(angle) & " zigzag=" & CStr(pattern = 1)
            AAProgressCandidate
            AADebugCandidateStart candidateLabel
            buildError = vbNullString
            On Error Resume Next
            Set candidate = BuildPass(models, designIndexes, designCount, cutIndexes, _
                cutCount, quantities, settings, forbidden, CDbl(angle), (pattern = 1), singleAreaOnly)
            If Err.Number <> 0 Then
                buildError = "Error " & CStr(Err.Number) & ": " & Err.Description
                lastError = Err.Description
            End If
            Err.Clear
            On Error GoTo 0
            AADebugCandidateEnd
            If Len(buildError) > 0 Then AADebugWrite "Medium", candidateLabel & " -> " & buildError
            Set candidate = FinalizeCandidate(candidate, models, forbidden, settings, _
                sourceDocument, sharedPlanner, candidateLabel)
            ConsiderCandidate candidate, best, candidateLabel, bestLabel, sharedPlanner
            If capacityTarget > 0 Then
                If Not candidate Is Nothing Then
                    If candidate.MatchesQuotas(designIndexes, designCount, quantities, capacityTarget) Then
                        ' Validasi posisi sesudah centering/repair sebelum shortcut kapasitas.
                        If ValidPlan(candidate, forbidden, settings, "Capacity") Then
                            Set best = candidate: bestLabel = candidateLabel
                            capacityValidated = True
                            GoTo BuildComplete
                        End If
                    End If
                End If
            End If
    NextPattern:
        Next pattern
    Next angle
    If best Is Nothing Then Err.Raise vbObjectError + 5241, "AAMediumPlanner", _
        "Tidak ditemukan susunan Medium yang memenuhi area, gap, dan mark." & _
        IIf(Len(lastError) > 0, vbCrLf & lastError, vbNullString)
    If capacityTarget > 0 Then
        If best.DesignCount < capacityTarget Then
            ' Compact mempertahankan jumlah; tidak dapat memenuhi quota yang kurang.
            AADebugWrite "Capacity", "target=" & CStr(capacityTarget) & "; partial=" & _
                CStr(best.DesignCount) & "; Compact skipped"
            GoTo BuildComplete
        End If
    End If
    If best.HasRegularPattern Then
        ' Compaction tidak menambah objek dan dapat merusak pola teratur.
        AADebugWrite "Compact", "skipped: selected plan has a regular pattern"
    Else
        candidateLabel = "Compact " & bestLabel
        AAProgressCandidate
        AADebugCandidateStart candidateLabel
        AAProgressStage "Merapatkan susunan"
        Set candidate = CompactPlan(best, models, forbidden, settings)
        AADebugCandidateEnd
        Set candidate = FinalizeCandidate(candidate, models, forbidden, settings, _
            sourceDocument, sharedPlanner, candidateLabel)
        If Not candidate Is Nothing Then
            If ValidPlan(candidate, forbidden, settings, "Compact final") Then
                If capacityTarget > 0 Then
                    If candidate.MatchesQuotas(designIndexes, designCount, quantities, capacityTarget) Then
                        Set best = candidate: bestLabel = candidateLabel
                        capacityValidated = True
                        GoTo BuildComplete
                    End If
                End If
                ConsiderCandidate candidate, best, candidateLabel, bestLabel, sharedPlanner
            Else
                AADebugWrite "Compact", "rejected: final validation"
            End If
        End If
    End If
    ' Kuota sudah ditentukan. Alignment hanya memperbaiki hasil akhir, bukan trial kapasitas.
    If capacityTarget < 0 And Not best.HasRegularPattern Then
        candidateLabel = "Align neighbors " & bestLabel
        AAProgressCandidate
        AADebugCandidateStart candidateLabel
        AAProgressStage "Menyelaraskan susunan"
        Set candidate = AlignNeighborPlan(best, models, forbidden, settings)
        AADebugCandidateEnd
        Set candidate = FinalizeCandidate(candidate, models, forbidden, settings, _
            sourceDocument, sharedPlanner, candidateLabel)
        If Not candidate Is Nothing Then
            If ValidPlan(candidate, forbidden, settings, "Alignment final") Then
                ' Gerak lokal tidak memperluas envelope. Setelah centering, jangan
                ' mengganti hasil jika pembulatan justru memperburuk skor planner.
                If Not sharedPlanner.BetterPlan(best, candidate) Then
                    Set best = candidate: bestLabel = candidateLabel
                    AADebugWrite "Align", "accepted after final validation"
                Else
                    AADebugWrite "Align", "retained previous plan: final bounds score"
                End If
            Else
                AADebugWrite "Align", "rejected: final validation"
            End If
        End If
    End If
BuildComplete:
    If capacityTarget > 0 Then
        If best.MatchesQuotas(designIndexes, designCount, quantities, capacityTarget) Then
            If Not capacityValidated Then Err.Raise vbObjectError + 5241, "AAMediumPlanner", _
                "Quota penuh tidak lolos validasi posisi akhir."
        End If
        AADebugWrite "Capacity", "target=" & CStr(capacityTarget) & "; selected=" & bestLabel & _
            "; complete=" & CStr(best.MatchesQuotas(designIndexes, designCount, quantities, capacityTarget))
    End If
    AADebugPlan "Selected " & bestLabel, best
    If Not mQuotaFills Is Nothing Then AADebugWrite "Quota cache", _
        "hits=" & CStr(mFillHits) & "; misses=" & CStr(mFillMisses) & _
        "; reused placements=" & CStr(mFillReused) & "; stored=" & CStr(mCachedPlacements)
    If Not mQuotaFills Is Nothing Then AADebugWrite "Quota gap cache", _
        "hits=" & Format$(mGapHits, "0") & "; misses=" & Format$(mGapMisses, "0") & _
        "; collisions=" & Format$(mGapCollisions, "0") & "; bypass=" & Format$(mGapBypasses, "0") & _
        "; stored=" & CStr(mGapStored) & "; slots=" & CStr(AA_QUOTA_GAP_CACHE_SIZE) & _
        "; buckets=" & CStr(AA_QUOTA_GAP_BUCKETS) & "; ways=" & CStr(AA_QUOTA_GAP_WAYS)
    Set mOccupiedIndex = Nothing
    Set Build = best
End Function

Private Sub ConsiderCandidate(ByVal candidate As AALayoutPlan, ByRef best As AALayoutPlan, _
                              ByVal label As String, ByRef bestLabel As String, _
                              ByVal sharedPlanner As AAPlanner)
    Dim selected As Boolean
    If candidate Is Nothing Then Exit Sub
    AADebugPlan label, candidate
    If best Is Nothing Then
        selected = True
    Else
        selected = sharedPlanner.BetterPlan(candidate, best)
    End If
    If selected Then
        AADebugWrite "Plan", "Selected " & label & " (previous=" & bestLabel & ")"
        Set best = candidate
        bestLabel = label
    End If
End Sub

Private Function FinalizeCandidate(ByVal candidate As AALayoutPlan, ByRef models() As AAModel, _
                                   ByVal forbidden As Collection, ByVal settings As AASettings, _
                                   ByVal sourceDocument As Document, ByVal sharedPlanner As AAPlanner, _
                                   ByVal label As String) As AALayoutPlan
    Dim perfStarted As Double
    AADebugPlan label & " raw", candidate
    If candidate Is Nothing Then Exit Function
    ' FinalizeMedium memindahkan placement; indeks yang memegangnya harus dibatalkan.
    If Not mOccupiedIndex Is Nothing Then mOccupiedIndex.Clear
    perfStarted = AADebugPerfBegin()
    If Not sharedPlanner.FinalizeMedium(candidate, settings, sourceDocument) Then
        AADebugWrite "Center", label & " -> individual repair required: area or mark"
        AADebugCandidateStart label & " repair"
        Set candidate = RepairCentered(candidate, models, forbidden, settings)
        AADebugCandidateEnd
        If candidate Is Nothing Then AADebugWrite "Center", label & " -> rejected: repair failed"
    Else
        AADebugWrite "Center", label & " -> OK (common shift)"
    End If
    AADebugStageEnd "finalize: " & label, perfStarted
    Set FinalizeCandidate = candidate
End Function

Private Function CompactPlan(ByVal plan As AALayoutPlan, ByRef models() As AAModel, _
                             ByVal forbidden As Collection, _
                             ByVal settings As AASettings) As AALayoutPlan
    Dim current() As AAPlacement, placed As AAPlacement, occupied As Collection
    Dim orientation As AARotationVariant, result As AALayoutPlan
    Dim i As Long, j As Long, pass As Long, probes As Long, areaIndex As Long
    Dim visit As Long, skipped As Long, totalProbes As Long
    Dim hasCurve As Boolean
    Dim moved As Long, previousX As Double, previousY As Double
    For Each placed In plan.Placements
        If Not models(placed.DesignIndex).IsRectangle Then hasCurve = True
    Next placed
    If Not hasCurve Then Exit Function
    AADebugWrite "Compact", "clearance reserve=" & Format$(AA_COMPACTION_RESERVE_MM, "0.000") & _
        " mm; copied gap validation=True"
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
        moved = 0: skipped = 0: totalProbes = 0
        ' Pass kedua memberi objek akhir kesempatan sebelum anggaran area habis.
        For visit = 1 To plan.DesignCount
            AAProgressPulse
            i = visit
            If pass Mod 2 = 0 Then i = plan.DesignCount - visit + 1
            If current(i).AreaIndex <> areaIndex Then
                areaIndex = current(i).AreaIndex
                probes = 0
            End If
            ' Rectangle mempertahankan anchor remainder Align Top / Align Left.
            If models(current(i).DesignIndex).IsRectangle Then GoTo NextPlaced
            If probes >= AA_MAX_COMPACTION_PROBES_PER_AREA Then
                skipped = skipped + 1
                GoTo NextPlaced
            End If
            totalProbes = totalProbes - probes
            Set occupied = New Collection
            For j = 1 To plan.DesignCount
                If j <> i And current(j).AreaIndex = areaIndex Then occupied.Add current(j)
            Next j
            Set orientation = models(current(i).DesignIndex).RotationVariant(current(i).RotationDelta)
            previousX = current(i).Footprint.LeftX: previousY = current(i).Footprint.TopY
            Set current(i) = CompactAxis(current(i), orientation, occupied, forbidden, settings, True, probes)
            Set current(i) = CompactAxis(current(i), orientation, occupied, forbidden, settings, False, probes)
            totalProbes = totalProbes + probes
            If Abs(current(i).Footprint.LeftX - previousX) > AA_EDGE_TOLERANCE_MM Or _
               Abs(current(i).Footprint.TopY - previousY) > AA_EDGE_TOLERANCE_MM Then moved = moved + 1
NextPlaced:
        Next visit
        AADebugWrite "Compact", "pass=" & CStr(pass) & " moved=" & CStr(moved) & _
            "; order=" & IIf(pass Mod 2 = 0, "reverse", "forward") & _
            "; probes=" & CStr(totalProbes) & "; skipped budget=" & CStr(skipped)
    Next pass
    Set result = New AALayoutPlan
    result.Configure AA_MODE_MEDIUM
    For i = 1 To plan.DesignCount
        result.Add current(i)
    Next i
    Set CompactPlan = result
End Function

Private Function AlignNeighborPlan(ByVal plan As AALayoutPlan, ByRef models() As AAModel, _
                                   ByVal forbidden As Collection, ByVal settings As AASettings) As AALayoutPlan
    Dim current() As AAPlacement, occupied As Collection, result As AALayoutPlan
    Dim candidate As AAPlacement, orientation As AARotationVariant
    Dim probes() As Long, i As Long, j As Long, pass As Long, visit As Long
    Dim moved As Long, totalMoved As Long, areaIndex As Long
    Dim x As Double, y As Double
    ReDim current(1 To plan.DesignCount)
    For i = 1 To plan.DesignCount
        Set candidate = plan.Placements(i)
        Set current(i) = candidate.CopyForPlan()
    Next i
    ReDim probes(0 To plan.AreaCount - 1)
    For pass = 1 To AA_ALIGNMENT_PASSES
        moved = 0
        ' Kunjungi tetangga kanan/bawah lebih dulu, lalu ulangi untuk ruang yang baru terbuka.
        For visit = 1 To plan.DesignCount
            i = plan.DesignCount - visit + 1
            areaIndex = current(i).AreaIndex
            If Not models(current(i).DesignIndex).IsRectangle And _
               probes(areaIndex) < AA_MAX_ALIGNMENT_PROBES_PER_AREA Then
                Set occupied = New Collection
                For j = 1 To plan.DesignCount
                    If j <> i And current(j).AreaIndex = areaIndex Then occupied.Add current(j)
                Next j
                If NeighborAlignmentTarget(current(i), occupied, x, y) Then
                    probes(areaIndex) = probes(areaIndex) + 1
                    Set orientation = models(current(i).DesignIndex).RotationVariant(current(i).RotationDelta)
                    Set candidate = TryExactSlot(orientation, current(i).DesignIndex, current(i).CutIndex, _
                        areaIndex, x, y, occupied, forbidden, settings, current(i).GapH, current(i).GapV, _
                        AA_COMPACTION_RESERVE_MM)
                    If Not candidate Is Nothing Then
                        If AlignmentEnvelopeFits(current(i), candidate, occupied) Then
                            candidate.SetShiftOrigin current(i).SlotLeft, current(i).SlotTop
                            If Not current(i).CreatesCutLine Then candidate.SuppressCutLine
                            AADebugWrite "Align", "pass=" & CStr(pass) & "; object=" & CStr(i) & _
                                "; design=" & CStr(candidate.DesignIndex) & "; dx=" & _
                                Format$(candidate.Footprint.LeftX - current(i).Footprint.LeftX, "0.000000") & _
                                "; dy=" & Format$(candidate.Footprint.TopY - current(i).Footprint.TopY, "0.000000")
                            Set current(i) = candidate
                            moved = moved + 1
                        End If
                    End If
                End If
            End If
        Next visit
        totalMoved = totalMoved + moved
        AADebugWrite "Align", "pass=" & CStr(pass) & "; moved=" & CStr(moved)
        If moved = 0 Then Exit For
    Next pass
    For areaIndex = 0 To UBound(probes)
        AADebugWrite "Align", "area=" & CStr(areaIndex) & "; probes=" & CStr(probes(areaIndex)) & _
            "; limit=" & CStr(AA_MAX_ALIGNMENT_PROBES_PER_AREA)
    Next areaIndex
    If totalMoved = 0 Then Exit Function
    Set result = New AALayoutPlan
    result.Configure AA_MODE_MEDIUM
    For i = 1 To plan.DesignCount
        result.Add current(i)
    Next i
    Set AlignNeighborPlan = result
End Function

' Dua anchor lokal: center kolom dari atas, center baris dari kiri/kanan.
' Hanya source/rotasi yang sama; jangan menarik item ke baris/kolom yang jauh.
Private Function NeighborAlignmentTarget(ByVal placed As AAPlacement, ByVal occupied As Collection, _
                                         ByRef x As Double, ByRef y As Double) As Boolean
    Dim box As AAGeometry, other As AAGeometry, neighbor As AAPlacement
    Dim cx As Double, cy As Double, nx As Double, ny As Double, w As Double, h As Double
    Dim columnX As Double, rowY As Double, columnDistance As Double, rowDistance As Double
    Dim dx As Double, dy As Double, distance As Double, hasColumn As Boolean, hasRow As Boolean
    Set box = placed.CollisionFootprint
    w = box.RightX - box.LeftX: h = box.TopY - box.BottomY
    If w <= AA_EDGE_TOLERANCE_MM Or h <= AA_EDGE_TOLERANCE_MM Then Exit Function
    cx = (box.LeftX + box.RightX) / 2#: cy = (box.TopY + box.BottomY) / 2#
    For Each neighbor In occupied
        If neighbor.DesignIndex = placed.DesignIndex And neighbor.AreaIndex = placed.AreaIndex And _
           Abs(neighbor.RotationDelta - placed.RotationDelta) <= AA_EDGE_TOLERANCE_MM Then
            Set other = neighbor.CollisionFootprint
            nx = (other.LeftX + other.RightX) / 2#: ny = (other.TopY + other.BottomY) / 2#
            dx = Abs(nx - cx): dy = ny - cy
            If dx <= w * AA_SHIFT_FACTOR And dy > h * 0.5 And dy <= h * 2.5 Then
                distance = dy / h + dx / w
                If Not hasColumn Or distance < columnDistance Then
                    hasColumn = True: columnDistance = distance: columnX = nx
                End If
            End If
            If dx > w * 0.5 And dx <= w * 2# And Abs(dy) <= h * 0.4 Then
                distance = Abs(dy) / h + dx / w
                If Not hasRow Or distance < rowDistance Then
                    hasRow = True: rowDistance = distance: rowY = ny
                End If
            End If
        End If
    Next neighbor
    If Not hasColumn Or Not hasRow Then Exit Function
    dx = columnX - cx: dy = rowY - cy
    If Abs(dx) <= AA_COMPACTION_TOLERANCE_MM And Abs(dy) <= AA_COMPACTION_TOLERANCE_MM Then Exit Function
    x = placed.Footprint.LeftX + dx: y = placed.Footprint.TopY + dy
    NeighborAlignmentTarget = True
End Function

' Seluruh bbox artwork DAN collision harus tetap di dalam envelope sebelumnya.
' Gerak satu pasangan boleh turun/kanan tanpa memperlebar susunan area.
Private Function AlignmentEnvelopeFits(ByVal original As AAPlacement, ByVal candidate As AAPlacement, _
                                       ByVal occupied As Collection) As Boolean
    Dim box As AAGeometry, other As AAGeometry, neighbor As AAPlacement, kind As Long
    Dim leftX As Double, rightX As Double, topY As Double, bottomY As Double
    For kind = 0 To 1
        If kind = 0 Then Set box = original.Footprint Else Set box = original.CollisionFootprint
        leftX = box.LeftX: rightX = box.RightX: topY = box.TopY: bottomY = box.BottomY
        For Each neighbor In occupied
            If kind = 0 Then Set other = neighbor.Footprint Else Set other = neighbor.CollisionFootprint
            leftX = MinValue(leftX, other.LeftX): rightX = MaxValue(rightX, other.RightX)
            topY = MaxValue(topY, other.TopY): bottomY = MinValue(bottomY, other.BottomY)
        Next neighbor
        If kind = 0 Then Set box = candidate.Footprint Else Set box = candidate.CollisionFootprint
        If box.LeftX < leftX Or box.RightX > rightX Or box.TopY > topY Or box.BottomY < bottomY Then Exit Function
    Next kind
    AlignmentEnvelopeFits = True
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
    ' Compact mengikuti ruang menuju tetangga; slot awal tetap metadata asal.
    If horizontal Then
        maximum = placed.Footprint.LeftX
    Else
        maximum = settings.AreaHeightMM - placed.Footprint.TopY
    End If
    If maximum <= AA_COMPACTION_TOLERANCE_MM Then GoTo FinishedCompaction
    distance = MinValue(AA_SHIFT_STEP_MM, maximum)
    Do While probes < AA_MAX_COMPACTION_PROBES_PER_AREA
        x = placed.Footprint.LeftX: y = placed.Footprint.TopY
        If horizontal Then x = x - distance Else y = y + distance
        probes = probes + 1
        Set candidate = TryExactSlot(orientation, placed.DesignIndex, placed.CutIndex, _
            placed.AreaIndex, x, y, occupied, forbidden, settings, placed.GapH, placed.GapV, _
            AA_COMPACTION_RESERVE_MM)
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
            placed.AreaIndex, x, y, occupied, forbidden, settings, placed.GapH, placed.GapV, _
            AA_COMPACTION_RESERVE_MM)
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
    Dim moved As Long
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
            If candidate Is Nothing Then
                AADebugWrite "Repair", "failed at object=" & CStr(i) & "; moved=" & CStr(moved)
                Exit Function
            End If
            If Abs(candidate.Footprint.LeftX - current(i).Footprint.LeftX) > AA_EDGE_TOLERANCE_MM Or _
               Abs(candidate.Footprint.TopY - current(i).Footprint.TopY) > AA_EDGE_TOLERANCE_MM Then moved = moved + 1
            Set current(i) = candidate
        End If
    Next i
    Set result = New AALayoutPlan
    result.Configure AA_MODE_MEDIUM
    result.SetRegularPattern plan.HasRegularPattern And moved = 0
    For i = 1 To plan.DesignCount
        result.Add current(i)
    Next i
    Set RepairCentered = result
    AADebugWrite "Repair", "moved=" & CStr(moved) & "; regular=" & CStr(result.HasRegularPattern)
End Function

Private Function AdaptMinimum(ByVal source As AALayoutPlan, _
                              ByRef models() As AAModel) As AALayoutPlan
    Dim result As AALayoutPlan, placed As AAPlacement
    Dim orientation As AARotationVariant
    Dim h As Double, v As Double
    Dim firstIndex As Long, singleDesign As Boolean
    Set result = New AALayoutPlan
    result.Configure AA_MODE_MEDIUM
    firstIndex = -1
    singleDesign = True
    For Each placed In source.Placements
        If firstIndex < 0 Then firstIndex = placed.DesignIndex
        If placed.DesignIndex <> firstIndex Then singleDesign = False
        models(placed.DesignIndex).GetBoundaryGaps AA_MODE_STRAIGHT, h, v
        Set orientation = models(placed.DesignIndex).RotationVariant(0#)
        result.Add orientation.MakePlacement(placed.DesignIndex, placed.CutIndex, _
            placed.AreaIndex, placed.Footprint.LeftX, placed.Footprint.TopY, h, v)
    Next placed
    result.SetRegularPattern source.DesignCount > 0 And singleDesign
    Set AdaptMinimum = result
End Function

Private Function ValidPlan(ByVal plan As AALayoutPlan, ByVal forbidden As Collection, _
                           ByVal settings As AASettings, _
                           Optional ByVal validationLabel As String = "Plan") As Boolean
    Dim occupied As Collection, placed As AAPlacement
    Dim areaIndex As Long, objectIndex As Long
    Set occupied = New Collection
    areaIndex = -1
    For Each placed In plan.Placements
        If placed.AreaIndex <> areaIndex Then
            areaIndex = placed.AreaIndex
            Set occupied = New Collection
        End If
        objectIndex = objectIndex + 1
        If Not Fits(placed, occupied, forbidden, settings, False, _
            validationLabel & "; object=" & CStr(objectIndex)) Then Exit Function
        occupied.Add placed
    Next placed
    ValidPlan = True
End Function

Private Function BuildPass(ByRef models() As AAModel, ByRef designIndexes() As Long, _
                           ByVal designCount As Long, ByRef cutIndexes() As Long, _
                           ByVal cutCount As Long, ByRef quantities() As Long, _
                           ByVal settings As AASettings, ByVal forbidden As Collection, _
                           ByVal firstAngle As Double, ByVal zigZag As Boolean, _
                           ByVal singleAreaOnly As Boolean) As AALayoutPlan
    Dim result As AALayoutPlan, occupied As Collection
    Dim roleIndex As Long, index As Long, cutIndex As Long, areaIndex As Long
    Dim remaining As Long, added As Long, sourceCount As Long, requested As Boolean
    Dim prefix As String, fillKey As String
    prefix = CStr(firstAngle) & "|" & CStr(zigZag)
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
            If singleAreaOnly Then GoTo NextSource
            If added = 0 And occupied.Count > 0 Then
                areaIndex = areaIndex + 1
                Set occupied = New Collection
                added = FillArea(result, models(index), index, cutIndex, areaIndex, _
                    occupied, forbidden, settings, AA_MAX_OBJECTS_PER_SOURCE, firstAngle, zigZag)
            End If
            If added = 0 Then Err.Raise vbObjectError + 5231, _
                "AAMediumPlanner", "Object tidak muat dalam satu area pada pencarian Medium."
        Else
            sourceCount = 0
            Do While remaining > 0
                If singleAreaOnly And Not mQuotaFills Is Nothing Then
                    fillKey = prefix & "|" & CStr(index) & ":" & CStr(cutIndex)
                    added = FillQuotaArea(fillKey, result, models(index), index, cutIndex, _
                        occupied, forbidden, settings, remaining, firstAngle, zigZag)
                Else
                    added = FillArea(result, models(index), index, cutIndex, areaIndex, _
                        occupied, forbidden, settings, remaining, firstAngle, zigZag)
                End If
                If added = 0 And occupied.Count = 0 Then Err.Raise vbObjectError + 5231, _
                    "AAMediumPlanner", "Object tidak muat dalam satu area pada pencarian Medium."
                remaining = remaining - added
                If singleAreaOnly And remaining > 0 Then Exit For
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
        ' Quota seluruh sumber sebelumnya menentukan occupied untuk sumber ini.
        ' Quantity 0 juga masuk key agar perpindahan peran tidak berbenturan.
        prefix = prefix & "|" & CStr(index) & ":" & CStr(cutIndex) & "=" & CStr(quantities(index))
    Next roleIndex
    If requested And result.DesignCount = 0 Then Err.Raise vbObjectError + 5233, _
        "AAMediumPlanner", "Tidak ada Design yang muat pada pencarian Medium."
    Set BuildPass = result
End Function

Private Function FillQuotaArea(ByVal key As String, ByVal result As AALayoutPlan, _
                               ByVal model As AAModel, ByVal designIndex As Long, _
                               ByVal cutIndex As Long, ByVal occupied As Collection, _
                               ByVal forbidden As Collection, ByVal settings As AASettings, _
                               ByVal limit As Long, ByVal firstAngle As Double, _
                               ByVal zigZag As Boolean) As Long
    Dim entry As Collection, saved As Collection, placed As AAPlacement
    Dim previousCount As Long, startCount As Long, added As Long, i As Long
    If mQuotaFills.Exists(key) Then
        Set entry = mQuotaFills(key)
        Set saved = entry(1)
        previousCount = saved.Count
        ' FillArea berhenti tepat pada limit. Prefix hanya dapat dipakai bila
        ' cukup panjang, atau pencarian lama sudah menghabiskan seluruh slot.
        If limit <= saved.Count Or CBool(entry(2)) Then
            added = saved.Count
            If added > limit Then added = limit
            For i = 1 To added
                ' Friend member harus dipanggil melalui reference AAPlacement,
                ' bukan Variant hasil Collection.Item (late binding).
                Set placed = saved(i)
                Set placed = placed.CopyForPlan()
                result.Add placed
                occupied.Add placed
            Next i
            mFillHits = mFillHits + 1
            mFillReused = mFillReused + added
            FillQuotaArea = added
            Exit Function
        End If
    End If
    mFillMisses = mFillMisses + 1
    startCount = result.DesignCount
    added = FillArea(result, model, designIndex, cutIndex, 0, occupied, forbidden, _
        settings, limit, firstAngle, zigZag)
    ' Error FillArea tidak masuk cache. Batas memori tidak memotong pencarian;
    ' entry yang tidak tertampung dihitung ulang dengan jalur biasa.
    If mCachedPlacements - previousCount + added <= AA_MAX_CACHED_PLACEMENTS Then
        Set saved = New Collection
        For i = 1 To added
            Set placed = result.Placements(startCount + i)
            saved.Add placed.CopyForPlan()
        Next i
        Set entry = New Collection
        entry.Add saved
        entry.Add (added < limit)
        If mQuotaFills.Exists(key) Then mQuotaFills.Remove key
        mQuotaFills.Add key, entry
        mCachedPlacements = mCachedPlacements - previousCount + added
    End If
    FillQuotaArea = added
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

Private Function CurvedPairEligible(ByVal orientation As AARotationVariant) As Boolean
    Dim nodes As Long, edges As Long
    nodes = orientation.Contour.Anchors.Count
    edges = orientation.Contour.EdgeCount
    ' Ellipse standar (4 node) sudah ditangani baseline/ZigZag. Lane baru
    ' hanya untuk kontur berlekuk sederhana; jumlah segmen tetap dibatasi.
    CurvedPairEligible = (nodes >= 5 And nodes <= AA_MAX_CURVED_PAIR_NODES And _
        edges > AA_MAX_PAIR_CONTOUR_EDGES And edges <= AA_MAX_CURVED_PAIR_EDGES)
End Function

Private Sub ConsiderCurvedPairs(ByVal model As AAModel, ByVal designIndex As Long, _
                                ByVal cutIndex As Long, ByVal settings As AASettings, _
                                ByVal forbidden As Collection, ByRef models() As AAModel, _
                                ByVal sourceDocument As Document, ByVal sharedPlanner As AAPlanner, _
                                ByRef best As AALayoutPlan, ByRef bestLabel As String)
    Dim orientation As AARotationVariant, opposite As AARotationVariant
    Dim candidate As AALayoutPlan
    Dim angle As Long, phaseIndex As Long, rowMode As Long
    Dim phaseLimit As Double, phaseY As Double, h As Double, v As Double
    Dim label As String, buildError As String

    model.GetBoundaryGaps AA_MODE_STRAIGHT, h, v
    If h < 0# Or v < 0# Then Exit Sub
    ' Dua keluarga: 0/180 dan 90/270. Sembilan fase termasuk kedua tanda,
    ' dua offset baris: maksimal 36 kandidat; tidak mengikuti trial quota.
    For angle = 0 To 90 Step 90
        Set orientation = model.RotationVariant(CDbl(angle))
        Set opposite = model.RotationVariant(CDbl(angle + 180))
        If Not CurvedPairEligible(orientation) Or Not CurvedPairEligible(opposite) Then GoTo NextCurvedAngle
        phaseLimit = MinValue(orientation.ReferenceHeight, opposite.ReferenceHeight) * AA_SHIFT_FACTOR
        For phaseIndex = -AA_CURVED_PHASE_STEPS To AA_CURVED_PHASE_STEPS
            phaseY = phaseLimit * phaseIndex / AA_CURVED_PHASE_STEPS
            For rowMode = 0 To 1
                AAProgressCandidate
                label = "Curved pair angle=" & CStr(angle) & " row=" & CStr(rowMode) & _
                    " phase=" & Format$(phaseY, "0.000000")
                AADebugCandidateStart label
                Set candidate = Nothing
                buildError = vbNullString
                On Error Resume Next
                Set candidate = BuildOppositeLattice(model, designIndex, cutIndex, settings, _
                    forbidden, CDbl(angle), rowMode, 4, phaseY)
                If Err.Number <> 0 Then buildError = "Error " & CStr(Err.Number) & ": " & Err.Description
                Err.Clear
                On Error GoTo 0
                AADebugCandidateEnd
                If Len(buildError) > 0 Then AADebugWrite "Pair", label & " -> " & buildError
                Set candidate = FinalizeCandidate(candidate, models, forbidden, settings, _
                    sourceDocument, sharedPlanner, label)
                If Not candidate Is Nothing Then
                    If ValidPlan(candidate, forbidden, settings, label) Then _
                        ConsiderCandidate candidate, best, label, bestLabel, sharedPlanner
                End If
            Next rowMode
        Next phaseIndex
NextCurvedAngle:
    Next angle
End Sub

Private Function BuildOppositeLattice(ByVal model As AAModel, _
                                      ByVal designIndex As Long, ByVal cutIndex As Long, _
                                      ByVal settings As AASettings, _
                                      ByVal forbidden As Collection, _
                                      ByVal firstAngle As Double, _
                                      ByVal rowMode As Long, _
                                      ByVal patternMode As Long, _
                                      Optional ByVal phaseOffset As Double = 0#) As AALayoutPlan
    Dim forward As AARotationVariant, reverse As AARotationVariant
    Dim result As AALayoutPlan, occupied As Collection
    Dim orientation As AARotationVariant, candidate As AAPlacement
    Dim h As Double, v As Double, stepX As Double, stepY As Double
    Dim x As Double, y As Double, rowShift As Double, candidateTop As Double
    Dim rowIndex As Long, columnIndex As Long, probes As Long
    Dim scanContacts As Boolean, phaseY As Double
    Dim slotLimit As Long
    Dim startX As Double, spanX As Double, spanY As Double, rows As Long, columns As Long

    model.GetBoundaryGaps AA_MODE_STRAIGHT, h, v
    If h < 0# Or v < 0# Then Exit Function
    Set forward = model.RotationVariant(firstAngle)
    Set reverse = model.RotationVariant((firstAngle + 180#) Mod 360#)
    slotLimit = AA_MAX_SLOTS_PER_AREA
    If patternMode = 1 Or patternMode = 3 Then
        ' Pasangan kedua naik sebesar Gap V, dalam batas pergeseran Medium.
        phaseY = MinValue(v, MinValue(forward.ReferenceHeight, _
            reverse.ReferenceHeight) * AA_SHIFT_FACTOR)
        If phaseY <= AA_EDGE_TOLERANCE_MM Then Exit Function
        slotLimit = AA_MAX_PHASE_SLOTS_PER_AREA
    End If
    scanContacts = Not model.IsTriangle
    If patternMode = 4 Then
        If Not CurvedPairEligible(forward) Or Not CurvedPairEligible(reverse) Then Exit Function
        If Abs(phaseOffset) > MinValue(forward.ReferenceHeight, reverse.ReferenceHeight) * _
            AA_SHIFT_FACTOR + AA_EDGE_TOLERANCE_MM Then Exit Function
        phaseY = phaseOffset
        slotLimit = AA_MAX_PHASE_SLOTS_PER_AREA
        ' Lane ini memakai bracket aman + bisection; validasi detail di TryExactSlot
        ' dan ValidPlan tetap wajib. Tidak menjalankan 64 contact scan per pitch.
        scanContacts = False
    ElseIf scanContacts Then
        If forward.Contour.EdgeCount > AA_MAX_PAIR_CONTOUR_EDGES Then Exit Function
    End If
    If patternMode = 2 Or patternMode = 3 Then
        Set BuildOppositeLattice = BuildTightPairLattice(forward, reverse, _
            designIndex, cutIndex, settings, forbidden, h, v, phaseY, scanContacts)
        Exit Function
    End If
    stepX = OppositeHorizontalPitch(forward, reverse, h, v, phaseY, scanContacts)
    If stepX <= AA_COMPACTION_TOLERANCE_MM Then Exit Function
    If patternMode = 4 Then stepX = stepX + AA_PAIR_INSET_RESERVE_MM
    If rowMode = 1 Then rowShift = stepX / 2#
    stepY = OppositeVerticalPitch(forward, reverse, stepX, rowShift, _
        h, v, phaseY, scanContacts)
    If stepY <= AA_COMPACTION_TOLERANCE_MM Then Exit Function
    If patternMode = 4 Then stepY = stepY + AA_PAIR_INSET_RESERVE_MM
    Set result = New AALayoutPlan
    result.Configure AA_MODE_MEDIUM
    Set occupied = New Collection
    y = settings.AreaHeightMM - MaxValue(0#, phaseY)
    If patternMode = 4 Then
        ' Pusatkan grid SEBELUM filter mark. Centering sesudah filter tidak
        ' mengembalikan objek sudut yang sudah ditolak pada anchor (0, top).
        columns = Int((settings.AreaWidthMM - MaxValue(forward.Width, reverse.Width)) / stepX) + 1
        rows = Int((settings.AreaHeightMM - MaxValue(forward.Height, reverse.Height) - Abs(phaseY)) / stepY) + 1
        If columns > 0 Then
            spanX = (columns - 1) * stepX + MaxValue(forward.Width, reverse.Width)
            startX = (settings.AreaWidthMM - spanX) / 2#
        End If
        If rows > 0 Then
            spanY = (rows - 1) * stepY + MaxValue(forward.Height, reverse.Height) + Abs(phaseY)
            y = (settings.AreaHeightMM + spanY) / 2# - MaxValue(0#, phaseY)
        End If
    End If
    result.SetRegularPattern True
    If AADebugEnabled Then AADebugWrite "Pair", "angle=" & CStr(firstAngle) & _
        " row=" & CStr(rowMode) & " pattern=" & CStr(patternMode) & _
        "; phase=" & Format$(phaseY, "0.000000") & "; step=" & _
        Format$(stepX, "0.000000") & "/" & Format$(stepY, "0.000000")
    Do While y >= MinValue(forward.Height, reverse.Height - phaseY) - _
        AA_EDGE_TOLERANCE_MM
        x = startX
        If rowIndex Mod 2 = 1 Then x = x + rowShift
        columnIndex = 0
        Do While x <= settings.AreaWidthMM + AA_EDGE_TOLERANCE_MM
            probes = probes + 1
            If probes > slotLimit Then Exit Do
            If columnIndex Mod 2 = 0 Then
                Set orientation = forward
            Else
                Set orientation = reverse
            End If
            candidateTop = y
            If columnIndex Mod 2 = 1 Then candidateTop = candidateTop + phaseY
            Set candidate = TryExactSlot(orientation, designIndex, cutIndex, 0, _
                x, candidateTop, occupied, forbidden, settings, h, v)
            If Not candidate Is Nothing Then
                result.Add candidate
                occupied.Add candidate
                If result.DesignCount >= AA_MAX_OBJECTS_PER_SOURCE Then Exit Do
            End If
            x = x + stepX
            columnIndex = columnIndex + 1
        Loop
        If result.DesignCount >= AA_MAX_OBJECTS_PER_SOURCE Or _
           probes > slotLimit Then Exit Do
        y = y - stepY
        rowIndex = rowIndex + 1
    Loop
    If result.DesignCount > 0 Then Set BuildOppositeLattice = result
End Function

Private Function BuildTightPairLattice(ByVal forward As AARotationVariant, _
                                       ByVal reverse As AARotationVariant, _
                                       ByVal designIndex As Long, ByVal cutIndex As Long, _
                                       ByVal settings As AASettings, _
                                       ByVal forbidden As Collection, _
                                       ByVal h As Double, ByVal v As Double, _
                                       ByVal phaseY As Double, _
                                       ByVal scanContacts As Boolean) As AALayoutPlan
    Dim result As AALayoutPlan, occupied As Collection
    Dim candidate As AAPlacement
    Dim innerX As Double, pairStep As Double, rowStep As Double
    Dim x As Double, y As Double, probes As Long

    ' Dua objek satu pasangan: sela dalam mengikuti kontur,
    ' sedangkan pasangan berikutnya dipisahkan oleh bounds dan Gap H.
    innerX = OppositePairInset(forward, reverse, h, v, phaseY, scanContacts)
    If AADebugEnabled Then AADebugWrite "Pair", "angle=" & CStr(forward.Angle) & _
        "; phase=" & Format$(phaseY, "0.000000") & "; inset=" & Format$(innerX, "0.000000")
    If innerX < 0# Then Exit Function
    pairStep = MaxValue(forward.Width, innerX + reverse.Width) + h
    rowStep = MaxValue(forward.Height + phaseY, reverse.Height) + v
    If AADebugEnabled Then AADebugWrite "Pair", "pair step=" & Format$(pairStep, "0.000000") & _
        "; row step=" & Format$(rowStep, "0.000000")
    If pairStep <= AA_COMPACTION_TOLERANCE_MM Or _
       rowStep <= AA_COMPACTION_TOLERANCE_MM Then Exit Function
    Set result = New AALayoutPlan
    result.Configure AA_MODE_MEDIUM
    Set occupied = New Collection
    y = settings.AreaHeightMM - phaseY
    result.SetRegularPattern True
    Do While y >= MaxValue(forward.Height, reverse.Height - phaseY) - AA_EDGE_TOLERANCE_MM
        x = 0#
        Do While x <= settings.AreaWidthMM + AA_EDGE_TOLERANCE_MM
            probes = probes + 1
            If probes > AA_MAX_PHASE_SLOTS_PER_AREA Then Exit Do
            Set candidate = TryExactSlot(forward, designIndex, cutIndex, 0, _
                x, y, occupied, forbidden, settings, h, v)
            If Not candidate Is Nothing Then
                result.Add candidate
                occupied.Add candidate
                If result.DesignCount >= AA_MAX_OBJECTS_PER_SOURCE Then Exit Do
            End If
            probes = probes + 1
            If probes > AA_MAX_PHASE_SLOTS_PER_AREA Then Exit Do
            Set candidate = TryExactSlot(reverse, designIndex, cutIndex, 0, _
                x + innerX, y + phaseY, occupied, forbidden, settings, h, v)
            If Not candidate Is Nothing Then
                result.Add candidate
                occupied.Add candidate
                If result.DesignCount >= AA_MAX_OBJECTS_PER_SOURCE Then Exit Do
            End If
            x = x + pairStep
        Loop
        If result.DesignCount >= AA_MAX_OBJECTS_PER_SOURCE Or _
           probes > AA_MAX_PHASE_SLOTS_PER_AREA Then Exit Do
        y = y - rowStep
    Loop
    If result.DesignCount > 0 Then Set BuildTightPairLattice = result
End Function

Private Function OppositePairInset(ByVal forward As AARotationVariant, _
                                   ByVal reverse As AARotationVariant, _
                                   ByVal h As Double, ByVal v As Double, _
                                   ByVal phaseY As Double, _
                                   ByVal scanContacts As Boolean) As Double
    Dim first As AAPlacement, second As AAPlacement
    Dim low As Double, high As Double, middle As Double, topY As Double
    Dim scanUpper As Double, trial As Double
    Dim i As Long

    topY = MaxValue(forward.Height, reverse.Height)
    high = MaxValue(forward.Width, reverse.Width) + MaxValue(h, v) + 2#
    Set first = forward.MakePlacement(0, -1, 0, 0#, topY, h, v)
    Set second = reverse.MakePlacement(0, -1, 0, 0#, topY + phaseY, h, v)
    If OppositeClear(first, second) Then Exit Function
    Set second = reverse.MakePlacement(0, -1, 0, high, topY + phaseY, h, v)
    If Not OppositeClear(first, second) Then
        OppositePairInset = -1#
        Exit Function
    End If
    ' Kurva cekung dapat punya beberapa rentang kontak aman.
    If scanContacts Then
        scanUpper = high
        For i = 1 To AA_PAIR_CONTACT_SCAN_STEPS
            trial = scanUpper * i / AA_PAIR_CONTACT_SCAN_STEPS
            Set second = reverse.MakePlacement(0, -1, 0, trial, topY + phaseY, h, v)
            If OppositeClear(first, second) Then
                low = scanUpper * (i - 1) / AA_PAIR_CONTACT_SCAN_STEPS
                high = trial
                Exit For
            End If
        Next i
    End If
    For i = 1 To 12
        middle = (low + high) / 2#
        Set second = reverse.MakePlacement(0, -1, 0, middle, topY + phaseY, h, v)
        If OppositeClear(first, second) Then
            high = middle
        Else
            low = middle
        End If
    Next i
    If scanContacts And (h > 0# Or v > 0#) Then
        ' Inset pada ambang toleransi dapat gagal setelah bounds kurva dibulatkan.
        ' Cadangan kecil menjauhkan kontak kurva custom dari ambang tersebut.
        ' Periksa ulang karena kurva cekung dapat punya rentang aman yang sempit.
        high = high + AA_PAIR_INSET_RESERVE_MM
        Set second = reverse.MakePlacement(0, -1, 0, high, topY + phaseY, h, v)
        If Not OppositeClear(first, second) Then
            OppositePairInset = -1#
            Exit Function
        End If
    End If
    OppositePairInset = high
End Function

Private Function OppositeHorizontalPitch(ByVal forward As AARotationVariant, _
                                         ByVal reverse As AARotationVariant, _
                                         ByVal h As Double, ByVal v As Double, _
                                         ByVal phaseY As Double, _
                                         ByVal scanContacts As Boolean) As Double
    Dim low As Double, high As Double, middle As Double, topY As Double
    Dim scanUpper As Double, trial As Double, i As Long
    topY = MaxValue(forward.Height, reverse.Height) + Abs(phaseY)
    high = MaxValue(forward.Width, reverse.Width) + MaxValue(h, v) + 2#
    If Not OppositeRowClear(forward, reverse, high, topY, h, v, phaseY) Then Exit Function
    If scanContacts Then
        scanUpper = high
        For i = 1 To AA_PAIR_CONTACT_SCAN_STEPS
            trial = scanUpper * i / AA_PAIR_CONTACT_SCAN_STEPS
            If OppositeRowClear(forward, reverse, trial, topY, h, v, phaseY) Then
                low = scanUpper * (i - 1) / AA_PAIR_CONTACT_SCAN_STEPS
                high = trial
                Exit For
            End If
        Next i
    End If
    For i = 1 To 12
        middle = (low + high) / 2#
        If OppositeRowClear(forward, reverse, middle, topY, h, v, phaseY) Then
            high = middle
        Else
            low = middle
        End If
    Next i
    OppositeHorizontalPitch = high
End Function

Private Function OppositeRowClear(ByVal forward As AARotationVariant, _
                                  ByVal reverse As AARotationVariant, _
                                  ByVal stepX As Double, ByVal topY As Double, _
                                  ByVal h As Double, ByVal v As Double, _
                                  ByVal phaseY As Double) As Boolean
    Dim first As AAPlacement, second As AAPlacement, third As AAPlacement
    Set first = forward.MakePlacement(0, -1, 0, 0#, topY, h, v)
    Set second = reverse.MakePlacement(0, -1, 0, stepX, topY + phaseY, h, v)
    Set third = forward.MakePlacement(0, -1, 0, 2# * stepX, topY, h, v)
    If Not OppositeClear(first, second) Then Exit Function
    If Not OppositeClear(second, third) Then Exit Function
    OppositeRowClear = OppositeClear(first, third)
End Function

Private Function OppositeVerticalPitch(ByVal forward As AARotationVariant, _
                                       ByVal reverse As AARotationVariant, _
                                       ByVal stepX As Double, ByVal rowShift As Double, _
                                       ByVal h As Double, ByVal v As Double, _
                                       ByVal phaseY As Double, _
                                       ByVal scanContacts As Boolean) As Double
    Dim low As Double, high As Double, middle As Double
    Dim scanUpper As Double, trial As Double, i As Long
    high = MaxValue(forward.Height, reverse.Height) + Abs(phaseY) + MaxValue(h, v) + 2#
    If Not OppositeRowsClear(forward, reverse, stepX, rowShift, high, h, v, phaseY) Then Exit Function
    If scanContacts Then
        scanUpper = high
        For i = 1 To AA_PAIR_CONTACT_SCAN_STEPS
            trial = scanUpper * i / AA_PAIR_CONTACT_SCAN_STEPS
            If OppositeRowsClear(forward, reverse, stepX, rowShift, trial, h, v, phaseY) Then
                low = scanUpper * (i - 1) / AA_PAIR_CONTACT_SCAN_STEPS
                high = trial
                Exit For
            End If
        Next i
    End If
    For i = 1 To 12
        middle = (low + high) / 2#
        If OppositeRowsClear(forward, reverse, stepX, rowShift, middle, h, v, phaseY) Then
            high = middle
        Else
            low = middle
        End If
    Next i
    OppositeVerticalPitch = high
End Function

Private Function OppositeRowsClear(ByVal forward As AARotationVariant, _
                                   ByVal reverse As AARotationVariant, _
                                   ByVal stepX As Double, ByVal rowShift As Double, _
                                   ByVal stepY As Double, ByVal h As Double, _
                                   ByVal v As Double, ByVal phaseY As Double) As Boolean
    Dim upper(0 To 4) As AAPlacement, lower(0 To 4) As AAPlacement
    Dim orientation As AARotationVariant
    Dim topY As Double, candidateTop As Double, i As Long, j As Long
    topY = 2# * MaxValue(forward.Height, reverse.Height) + stepY
    For i = 0 To 4
        If i Mod 2 = 0 Then
            Set orientation = forward
        Else
            Set orientation = reverse
        End If
        candidateTop = topY
        If i Mod 2 = 1 Then candidateTop = candidateTop + phaseY
        Set upper(i) = orientation.MakePlacement(0, -1, 0, _
            i * stepX, candidateTop, h, v)
        Set lower(i) = orientation.MakePlacement(0, -1, 0, _
            rowShift + i * stepX, candidateTop - stepY, h, v)
    Next i
    For i = 0 To 4
        For j = 0 To 4
            If Not OppositeClear(upper(i), lower(j)) Then Exit Function
        Next j
    Next i
    OppositeRowsClear = True
End Function

Private Function OppositeClear(ByVal first As AAPlacement, _
                               ByVal second As AAPlacement) As Boolean
    If Not first.HasClearance(second) Then Exit Function
    OppositeClear = Not first.CollisionFootprint.Overlaps(second.CollisionFootprint)
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
                              ByVal h As Double, ByVal v As Double, _
                              Optional ByVal clearanceReserve As Double = 0#) As AAPlacement
    Dim candidate As AAPlacement
    AAProgressPulse
    If x < -AA_EDGE_TOLERANCE_MM Or y - orientation.Height < -AA_EDGE_TOLERANCE_MM Or _
       x + orientation.Width > settings.AreaWidthMM + AA_EDGE_TOLERANCE_MM Or _
       y > settings.AreaHeightMM + AA_EDGE_TOLERANCE_MM Then
        AADebugRejectSlot "bounds"
        Exit Function
    End If
    If Not HasClearanceAt(orientation, x, y, occupied, _
        h + clearanceReserve, v + clearanceReserve) Then Exit Function
    Set candidate = orientation.MakePlacement(designIndex, cutIndex, areaIndex, x, y, h, v)
    ' Compact memeriksa ulang gap pada bounds kurva yang sudah disalin/dibulatkan.
    If Fits(candidate, occupied, forbidden, settings, clearanceReserve = 0#) Then _
        Set TryExactSlot = candidate
End Function

' Tetangga disaring saja; semua predicate dan urutan pemeriksaan tetap di caller.
Private Function NearbyOccupied(ByVal occupied As Collection, ByVal leftX As Double, _
                                ByVal bottomY As Double, ByVal rightX As Double, _
                                ByVal topY As Double, ByVal gap As Double) As Collection
    If mOccupiedIndex Is Nothing Then Set mOccupiedIndex = New AAOccupiedIndex
    Set NearbyOccupied = mOccupiedIndex.Nearby(occupied, leftX, bottomY, rightX, topY, gap)
End Function

Private Function HasClearanceAt(ByVal orientation As AARotationVariant, ByVal x As Double, _
                                ByVal y As Double, ByVal occupied As Collection, _
                                ByVal h As Double, ByVal v As Double) As Boolean
    Dim placed As AAPlacement, other As AAGeometry, gap As Double
    Dim nearby As Collection, clear As Boolean
    Dim leftX As Double, rightX As Double, topY As Double, bottomY As Double
    leftX = x + orientation.InsetX: rightX = leftX + orientation.ReferenceWidth
    topY = y - orientation.InsetY: bottomY = topY - orientation.ReferenceHeight
    Set nearby = NearbyOccupied(occupied, leftX, bottomY, rightX, topY, MaxValue(h, v))
    For Each placed In nearby
        Set other = placed.CollisionFootprint
        gap = MaxValue(MaxValue(h, v), MaxValue(placed.GapH, placed.GapV))
        If leftX < other.RightX + gap And rightX > other.LeftX - gap And _
           bottomY < other.TopY + gap And topY > other.BottomY - gap Then
            If mQuotaFills Is Nothing Then
                clear = orientation.HasClearanceAt(x, y, placed, h, v)
            Else
                clear = QuotaClearanceAt(orientation, x, y, placed, h, v)
            End If
            If Not clear Then
                AADebugGapRejected orientation, x, y, placed, h, v
                Exit Function
            End If
        End If
    Next placed
    HasClearanceAt = True
End Function

' Cache hanya menahan kontur/orientasi milik models; tidak menahan placement
' yang dapat digeser. Key menyimpan snapshot posisi blocker dan gap efektif.
Private Function QuotaClearanceAt(ByVal orientation As AARotationVariant, _
                                  ByVal x As Double, ByVal y As Double, _
                                  ByVal placed As AAPlacement, ByVal h As Double, _
                                  ByVal v As Double) As Boolean
    Dim other As AAContour, box As AAGeometry, slot As Long
    Dim firstSlot As Long, emptySlot As Long, bucket As Long
    Dim gapH As Double, gapV As Double, clear As Boolean
    Set other = placed.Contour
    Set box = placed.CollisionFootprint
    gapH = MaxValue(h, placed.GapH): gapV = MaxValue(v, placed.GapV)
    firstSlot = QuotaGapSlot(x, y, box.LeftX, box.TopY, gapH, gapV, _
        orientation.Angle, placed.RotationDelta)
    If firstSlot < 0 Then
        mGapBypasses = mGapBypasses + 1#
        QuotaClearanceAt = orientation.HasClearanceAt(x, y, placed, h, v)
        Exit Function
    End If
    If Not mGapCacheReady Then
        If Not InitializeQuotaGapCache() Then
            mGapBypasses = mGapBypasses + 1#
            QuotaClearanceAt = orientation.HasClearanceAt(x, y, placed, h, v)
            Exit Function
        End If
    End If
    emptySlot = -1
    For slot = firstSlot To firstSlot + AA_QUOTA_GAP_WAYS - 1
        ' Slot baru selalu diisi berurutan; tidak ada lubang sebelum slot aktif.
        If mGapOrientations(slot) Is Nothing Then
            emptySlot = slot
            Exit For
        End If
        If mGapOrientations(slot) Is orientation And mGapOthers(slot) Is other Then
            If mGapKeys(0, slot) = x And mGapKeys(1, slot) = y And _
               mGapKeys(2, slot) = box.LeftX And mGapKeys(3, slot) = box.TopY And _
               mGapKeys(4, slot) = gapH And mGapKeys(5, slot) = gapV Then
                mGapHits = mGapHits + 1#
                QuotaClearanceAt = mGapResults(slot)
                Exit Function
            End If
        End If
    Next slot
    mGapMisses = mGapMisses + 1#
    bucket = firstSlot \ AA_QUOTA_GAP_WAYS
    If emptySlot >= 0 Then
        slot = emptySlot
    Else
        slot = firstSlot + mGapNextVictim(bucket)
        ' Collision menghitung penggantian bucket penuh, bukan slot yang diuji.
        mGapCollisions = mGapCollisions + 1#
    End If
    ' Simpan hanya sesudah predicate asli kembali normal; error tidak di-cache.
    clear = orientation.HasClearanceAt(x, y, placed, h, v)
    If mGapOrientations(slot) Is Nothing Then mGapStored = mGapStored + 1
    Set mGapOrientations(slot) = orientation: Set mGapOthers(slot) = other
    mGapKeys(0, slot) = x: mGapKeys(1, slot) = y
    mGapKeys(2, slot) = box.LeftX: mGapKeys(3, slot) = box.TopY
    mGapKeys(4, slot) = gapH: mGapKeys(5, slot) = gapV
    mGapResults(slot) = clear
    ' Penggantian bergilir hanya sesudah predicate sukses; hit/error tidak
    ' menggeser pointer. Lookup dibatasi empat slot, tanpa daftar tak terbatas.
    mGapNextVictim(bucket) = (slot - firstSlot + 1) Mod AA_QUOTA_GAP_WAYS
    QuotaClearanceAt = clear
End Function

Private Function InitializeQuotaGapCache() As Boolean
    Dim number As Long
    If mGapCacheDisabled Then Exit Function
    On Error GoTo CacheUnavailable
    ReDim mGapOrientations(0 To AA_QUOTA_GAP_CACHE_SIZE - 1)
    ReDim mGapOthers(0 To AA_QUOTA_GAP_CACHE_SIZE - 1)
    ReDim mGapKeys(0 To 5, 0 To AA_QUOTA_GAP_CACHE_SIZE - 1)
    ReDim mGapResults(0 To AA_QUOTA_GAP_CACHE_SIZE - 1)
    ReDim mGapNextVictim(0 To AA_QUOTA_GAP_BUCKETS - 1)
    mGapCacheReady = True
    InitializeQuotaGapCache = True
    Exit Function
CacheUnavailable:
    number = Err.Number
    On Error Resume Next
    Erase mGapOrientations: Erase mGapOthers: Erase mGapKeys: Erase mGapResults
    Erase mGapNextVictim
    mGapCacheReady = False: mGapCacheDisabled = True: mGapStored = 0
    AADebugWrite "Quota gap cache", "disabled: allocation error=" & CStr(number)
    Err.Clear
End Function

' Hash memilih slot pertama bucket. Hit membutuhkan identitas dan enam Double
' yang persis sama; tabrakan/koordinat besar kembali ke predicate biasa.
Private Function QuotaGapSlot(ByVal x As Double, ByVal y As Double, _
                              ByVal otherX As Double, ByVal otherY As Double, _
                              ByVal h As Double, ByVal v As Double, _
                              ByVal angle As Double, ByVal otherAngle As Double) As Long
    Dim key As Double
    QuotaGapSlot = -1
    If Abs(x) > AA_QUOTA_GAP_HASH_LIMIT Or Abs(y) > AA_QUOTA_GAP_HASH_LIMIT Or _
       Abs(otherX) > AA_QUOTA_GAP_HASH_LIMIT Or Abs(otherY) > AA_QUOTA_GAP_HASH_LIMIT Or _
       Abs(h) > AA_QUOTA_GAP_HASH_LIMIT Or Abs(v) > AA_QUOTA_GAP_HASH_LIMIT Or _
       Abs(angle) > AA_QUOTA_GAP_HASH_LIMIT Or Abs(otherAngle) > AA_QUOTA_GAP_HASH_LIMIT Then Exit Function
    ' Modulus prima menyebarkan koordinat grid integer yang berulang.
    ' Pada batas hash, jumlah integer berbobot ini tetap tepat dalam Double.
    key = CLng(Fix(x * 1024#)) * 31# + CLng(Fix(y * 1024#)) * 131# + _
        CLng(Fix(otherX * 1024#)) * 67# + CLng(Fix(otherY * 1024#)) * 257# + _
        CLng(Fix(h * 1024#)) * 17# + CLng(Fix(v * 1024#)) * 97# + _
        CLng(Fix(angle * 1024#)) * 193# + CLng(Fix(otherAngle * 1024#)) * 389#
    key = key - Fix(key / AA_QUOTA_GAP_BUCKETS) * AA_QUOTA_GAP_BUCKETS
    If key < 0# Then key = key + AA_QUOTA_GAP_BUCKETS
    QuotaGapSlot = CLng(key) * AA_QUOTA_GAP_WAYS
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
    If zigZag Then stepY = orientation.ReferenceHeight * settings.DefaultZigZagVerticalFactor
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
    Dim nearby As Collection
    Dim collisionLeft As Double, collisionTop As Double, gapH As Double, gapV As Double
    AAProgressPulse
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
    Set nearby = NearbyOccupied(occupied, collisionLeft + minX, _
        collisionTop - orientation.ReferenceHeight + minY, _
        collisionLeft + orientation.ReferenceWidth + maxX, collisionTop + maxY, MaxValue(h, v))
    For Each placed In nearby
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
            AAProgressPulse
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
                      Optional ByVal clearanceChecked As Boolean = False, _
                      Optional ByVal validationLabel As String = vbNullString) As Boolean
    Dim box As AAGeometry, collision As AAGeometry, other As AAGeometry
    Dim placed As AAPlacement, obstacle As AAGeometry, gap As Double
    Dim nearby As Collection
    Set box = candidate.Footprint
    If box.LeftX < -AA_EDGE_TOLERANCE_MM Or box.BottomY < -AA_EDGE_TOLERANCE_MM Or _
       box.RightX > settings.AreaWidthMM + AA_EDGE_TOLERANCE_MM Or _
       box.TopY > settings.AreaHeightMM + AA_EDGE_TOLERANCE_MM Then
        AADebugRejectSlot "bounds"
        If Len(validationLabel) > 0 Then AADebugValidationRejected _
            validationLabel, "bounds", candidate, Nothing
        Exit Function
    End If
    Set collision = candidate.CollisionFootprint
    For Each obstacle In forbidden
        If collision.Overlaps(obstacle) Then
            AADebugRejectSlot "mark"
            If Len(validationLabel) > 0 Then AADebugValidationRejected _
                validationLabel, "mark", candidate, Nothing
            Exit Function
        End If
    Next obstacle
    Set nearby = NearbyOccupied(occupied, collision.LeftX, collision.BottomY, _
        collision.RightX, collision.TopY, MaxValue(candidate.GapH, candidate.GapV))
    For Each placed In nearby
        Set other = placed.CollisionFootprint
        gap = MaxValue(MaxValue(candidate.GapH, candidate.GapV), MaxValue(placed.GapH, placed.GapV))
        If collision.LeftX < other.RightX + gap And collision.RightX > other.LeftX - gap And _
           collision.BottomY < other.TopY + gap And collision.TopY > other.BottomY - gap Then
            If Not clearanceChecked Then
                If Not candidate.HasClearance(placed) Then
                    AADebugRejectSlot "gap"
                    If Len(validationLabel) > 0 Then AADebugValidationRejected _
                        validationLabel, "gap", candidate, placed
                    Exit Function
                End If
            End If
            If collision.Overlaps(other) Then
                AADebugOverlapRejected candidate, placed
                If Len(validationLabel) > 0 Then AADebugValidationRejected _
                    validationLabel, "overlap", candidate, placed
                Exit Function
            End If
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
