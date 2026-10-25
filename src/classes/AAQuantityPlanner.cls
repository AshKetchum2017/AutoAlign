Option Explicit

Private Const AA_MAX_OBJECTS_PER_SOURCE As Long = 5000
Private Const AA_QUANTITY_MAX_DIFFERENCE As Long = 2
Private mSharedTemplate As AALayoutPlan
Private mTemplateCutIndex As Long
Private mTemplateRotationMask As Long
Private mQuotaPlanner As AAMediumPlanner
Private mLastCapacityCount As Long

' Satu container: target tetap menurut urutan, lalu filler seimbang.
' Target yang belum muat dikembalikan sebagai hasil parsial untuk scheduler.
Public Function BuildAllocatedGroup(ByRef models() As AAModel, ByRef designIndexes() As Long, _
                                    ByVal designCount As Long, ByRef cutIndexes() As Long, _
                                    ByVal cutCount As Long, ByRef requested() As Long, _
                                    ByRef limits() As Long, ByVal settings As AASettings, _
                                    ByVal sourceDocument As Document, ByVal layoutMode As Long, _
                                    ByVal planner As AAPlanner) As AALayoutPlan
    Dim result As AALayoutPlan, candidate As AALayoutPlan, pinned() As Long, trial() As Long
    Dim freeIndexes() As Long, freeCount As Long, index As Long, i As Long
    Dim low As Long, high As Long, middle As Long, expected As Long, room As Long, fillerCut As Long
    Dim fillerPlanner As AAMediumPlanner
    Dim allFree As Boolean, expanding As Boolean, needsFinalOptimization As Boolean
    ReDim pinned(0 To UBound(models)): ReDim trial(0 To UBound(models))
    ReDim freeIndexes(0 To designCount - 1)
    Set result = New AALayoutPlan: result.Configure layoutMode
    allFree = True
    For i = 0 To designCount - 1
        index = designIndexes(i)
        If requested(index) >= 0 Then allFree = False
        If requested(index) < 0 And limits(index) > 0 Then
            freeIndexes(freeCount) = index: freeCount = freeCount + 1
            If limits(index) <> AA_MAX_OBJECTS_PER_SOURCE Then allFree = False
        End If
    Next i
    If allFree And freeCount = designCount Then
        Set BuildAllocatedGroup = BuildGroup(models, designIndexes, designCount, cutIndexes, _
            cutCount, settings, sourceDocument, layoutMode, planner)
        Exit Function
    End If
    Set mQuotaPlanner = Nothing
    If CanUseTemplate(models, designIndexes, designCount, cutCount) Then
        Set candidate = AllocateTemplate(models, designIndexes, designCount, cutIndexes, cutCount, _
            requested, limits, settings, sourceDocument, layoutMode, planner)
        If Not candidate Is Nothing Then
            Set BuildAllocatedGroup = candidate
            Exit Function
        End If
    End If
    If layoutMode = AA_MODE_MEDIUM Then
        Set mQuotaPlanner = New AAMediumPlanner
        mQuotaPlanner.BeginQuotaSearch
    End If
    For i = 0 To designCount - 1
        AAProgressPulse
        index = designIndexes(i)
        If requested(index) > 0 Then
            CopyAllocation pinned, trial
            trial(index) = requested(index)
            If cutCount = 1 Then
                room = AA_MAX_OBJECTS_PER_SOURCE - result.DesignCount
                If trial(index) > room Then trial(index) = room
            End If
            If trial(index) > 0 Then
                Set candidate = PartialGroup(models, designIndexes, designCount, cutIndexes, cutCount, _
                    trial, settings, sourceDocument, layoutMode, planner)
                If Not candidate Is Nothing Then
                    If KeepsAllocation(candidate, pinned, index, trial(index)) Then
                        Set result = candidate
                        CountAllocation result, pinned
                    End If
                End If
            End If
        End If
    Next i
    If freeCount > 0 Then
        high = AA_MAX_OBJECTS_PER_SOURCE
        For i = 0 To freeCount - 1
            If limits(freeIndexes(i)) < high Then high = limits(freeIndexes(i))
        Next i
        If cutCount = 1 Then
            room = (AA_MAX_OBJECTS_PER_SOURCE - result.DesignCount) \ freeCount
            If room < high Then high = room
        End If
        expected = result.DesignCount
        CopyAllocation pinned, trial
        middle = 1: expanding = True
        Do While low < high
            For i = 0 To freeCount - 1
                trial(freeIndexes(i)) = middle
            Next i
            Set candidate = TryQuota(models, designIndexes, designCount, cutIndexes, cutCount, _
                trial, expected + middle * freeCount, settings, sourceDocument, layoutMode, planner)
            If candidate Is Nothing Then
                high = middle - 1: expanding = False
            Else
                low = middle: Set result = candidate
                needsFinalOptimization = True
            End If
            If expanding Then
                middle = low * 2
                If middle > high Then middle = high
            Else
                middle = (low + high + 1) \ 2
            End If
        Loop
        CopyAllocation pinned, trial
        For i = 0 To freeCount - 1
            trial(freeIndexes(i)) = low
        Next i
        expected = expected + low * freeCount
        ' Surplus satu per filler: target tetap tidak ikut diseimbangkan.
        For i = 0 To freeCount - 1
            index = freeIndexes(i)
            If trial(index) < limits(index) And (cutCount <> 1 Or expected < AA_MAX_OBJECTS_PER_SOURCE) Then
                trial(index) = trial(index) + 1
                Set candidate = TryQuota(models, designIndexes, designCount, cutIndexes, cutCount, _
                    trial, expected + 1, settings, sourceDocument, layoutMode, planner)
                If candidate Is Nothing Then
                    trial(index) = trial(index) - 1
                Else
                    expected = expected + 1: Set result = candidate
                    needsFinalOptimization = True
                End If
            End If
        Next i
        ' Jangan menerapkan hasil capacity-only mentah jika ada optimasi penuh
        ' dengan quota identik yang lebih baik. Rencana valid lama tetap cadangan.
        If needsFinalOptimization And layoutMode = AA_MODE_MEDIUM Then
            CountAllocation result, trial
            AADebugWrite "Quantity", "optimizing allocated quotas; target=" & CStr(result.DesignCount)
            Set candidate = TryQuota(models, designIndexes, designCount, cutIndexes, cutCount, _
                trial, result.DesignCount, settings, sourceDocument, layoutMode, planner, False)
            If Not candidate Is Nothing Then
                If planner.BetterPlan(candidate, result) Then Set result = candidate
            End If
        End If
        ' Isi celah pada posisi final, bukan seed quota sebelum centering/Compact.
        ' Satu filler tidak memerlukan pembagian seimbang antar beberapa sumber.
        If freeCount = 1 And cutCount <> 1 Then
            index = freeIndexes(0): fillerCut = -1
            For i = 0 To designCount - 1
                If designIndexes(i) = index And cutCount > 1 Then fillerCut = cutIndexes(i)
            Next i
            If layoutMode = AA_MODE_MEDIUM Then
                Set fillerPlanner = New AAMediumPlanner
                Set candidate = fillerPlanner.FillRemaining(result, models(index), index, fillerCut, _
                    limits(index), settings, sourceDocument, planner)
            Else
                Set candidate = planner.RefillMinimum(result, models(index), index, fillerCut, _
                    limits(index), settings, sourceDocument)
            End If
            If KeepsAllocation(candidate, pinned, index, limits(index)) And _
               candidate.DesignCount >= result.DesignCount Then Set result = candidate
        End If
    End If
    If result.AreaCount > 1 Or Not planner.FinalPlanFits(result, settings, sourceDocument) Then _
        Err.Raise vbObjectError + 5331, "AAQuantityPlanner", "Container Quantity tidak memenuhi area atau mark."
    Set mQuotaPlanner = Nothing
    Set BuildAllocatedGroup = result
End Function

Private Function AllocateTemplate(ByRef models() As AAModel, ByRef designs() As Long, _
                                   ByVal count As Long, ByRef cuts() As Long, ByVal cutCount As Long, _
                                   ByRef requested() As Long, ByRef limits() As Long, _
                                   ByVal settings As AASettings, ByVal doc As Document, _
                                   ByVal layoutMode As Long, ByVal planner As AAPlanner) As AALayoutPlan
    Dim template As AALayoutPlan, result As AALayoutPlan, source As AAPlacement
    Dim oneDesign(0 To 0) As Long, oneCut(0 To 0) As Long, quantities() As Long, allocated() As Long
    Dim i As Long, index As Long, slots As Long, amount As Long, slot As Long, j As Long, cut As Long
    Dim moved As Boolean
    oneDesign(0) = designs(0): ReDim quantities(0 To UBound(models))
    quantities(oneDesign(0)) = -1
    If cutCount > 0 Then oneCut(0) = cuts(0)
    If cutCount = 1 And Not mSharedTemplate Is Nothing Then
        If mTemplateCutIndex = cuts(0) And mTemplateRotationMask = models(designs(0)).RotationMask Then _
            Set template = mSharedTemplate
    End If
    If template Is Nothing Then
        Set template = PartialGroup(models, oneDesign, 1, oneCut, IIf(cutCount = 0, 0, 1), _
            quantities, settings, doc, layoutMode, planner)
        If template Is Nothing Then Exit Function
        If cutCount = 1 Then
            Set mSharedTemplate = template
            mTemplateCutIndex = cuts(0): mTemplateRotationMask = models(designs(0)).RotationMask
        End If
    End If
    ReDim allocated(0 To UBound(models)): slots = template.DesignCount
    For i = 0 To count - 1
        index = designs(i)
        If requested(index) > 0 Then
            amount = requested(index)
            If amount > slots Then amount = slots
            allocated(index) = amount: slots = slots - amount
        End If
    Next i
    Do While slots > 0
        moved = False
        For i = 0 To count - 1
            index = designs(i)
            If requested(index) < 0 And allocated(index) < limits(index) And slots > 0 Then
                allocated(index) = allocated(index) + 1: slots = slots - 1: moved = True
            End If
        Next i
        If Not moved Then Exit Do
        AAProgressPulse
    Loop
    Set result = New AALayoutPlan: result.Configure layoutMode
    result.SetRegularPattern template.HasRegularPattern: slot = 1
    For i = 0 To count - 1
        AAProgressPulse
        index = designs(i): cut = -1
        If cutCount = 1 Then cut = cuts(0)
        If cutCount > 1 Then cut = cuts(i)
        For j = 1 To allocated(index)
            AAProgressPulse
            Set source = template.Placements(slot)
            result.Add CopySlot(source, models(index), index, cut)
            slot = slot + 1
        Next j
    Next i
    If TemplateFits(result, planner, settings, doc) Then Set AllocateTemplate = result
End Function

Private Sub CopyAllocation(ByRef source() As Long, ByRef target() As Long)
    Dim i As Long
    For i = LBound(source) To UBound(source)
        target(i) = source(i)
    Next i
End Sub

Private Sub CountAllocation(ByVal plan As AALayoutPlan, ByRef counts() As Long)
    Dim i As Long, placed As AAPlacement
    For i = LBound(counts) To UBound(counts)
        counts(i) = 0
    Next i
    For Each placed In plan.Placements
        counts(placed.DesignIndex) = counts(placed.DesignIndex) + 1
    Next placed
End Sub

Private Function KeepsAllocation(ByVal plan As AALayoutPlan, ByRef pinned() As Long, _
                                 ByVal addedIndex As Long, ByVal maximum As Long) As Boolean
    Dim counts() As Long, i As Long
    If plan.AreaCount > 1 Then Exit Function
    ReDim counts(LBound(pinned) To UBound(pinned))
    CountAllocation plan, counts
    For i = LBound(pinned) To UBound(pinned)
        If i = addedIndex Then
            If counts(i) > maximum Then Exit Function
        ElseIf counts(i) <> pinned(i) Then
            Exit Function
        End If
    Next i
    KeepsAllocation = True
End Function

Private Function PartialGroup(ByRef models() As AAModel, ByRef designIndexes() As Long, _
                              ByVal designCount As Long, ByRef cutIndexes() As Long, _
                              ByVal cutCount As Long, ByRef quantities() As Long, _
                              ByVal settings As AASettings, ByVal sourceDocument As Document, _
                              ByVal layoutMode As Long, ByVal planner As AAPlanner) As AALayoutPlan
    Dim number As Long, source As String, description As String
    On Error GoTo Failed
    AAProgressTrial
    Set PartialGroup = BuildMode(models, designIndexes, designCount, cutIndexes, cutCount, _
        quantities, settings, sourceDocument, layoutMode, planner)
    Exit Function
Failed:
    number = Err.Number: source = Err.Source: description = Err.Description
    Select Case number
        Case vbObjectError + 5230, vbObjectError + 5231, vbObjectError + 5233, _
             vbObjectError + 5234, vbObjectError + 5241
            AADebugWrite "Quantity", "partial container rejected: " & description
        Case Else
            Err.Raise number, source, description
    End Select
End Function

' Satu instance per Process: cache hanya berlaku pada settings/dokumen yang sama.
Public Function BuildGroup(ByRef models() As AAModel, ByRef designIndexes() As Long, _
                           ByVal designCount As Long, ByRef cutIndexes() As Long, _
                           ByVal cutCount As Long, ByVal settings As AASettings, _
                           ByVal sourceDocument As Document, ByVal layoutMode As Long, _
                           ByVal planner As AAPlanner) As AALayoutPlan
    Dim quantities() As Long, oneDesign(0 To 0) As Long, oneCut(0 To 0) As Long
    Dim template As AALayoutPlan, result As AALayoutPlan, candidate As AALayoutPlan
    Dim low As Long, high As Long, middle As Long, i As Long, index As Long
    Dim expected As Long, pass As Long, failedAtTotal() As Long
    Dim hint As Long, consecutiveHints As Long
    Dim expanding As Boolean
    ' Geometri, urutan sumber, area dan gap cache tidak boleh lintas kelompok.
    Set mQuotaPlanner = Nothing
    ReDim quantities(0 To UBound(models))
    oneDesign(0) = designIndexes(0)
    quantities(oneDesign(0)) = -1
    If cutCount > 0 Then oneCut(0) = cutIndexes(0)

    If CanUseTemplate(models, designIndexes, designCount, cutCount) Then
        If cutCount = 1 Then
            If Not mSharedTemplate Is Nothing Then
                If mTemplateCutIndex = cutIndexes(0) And _
                   mTemplateRotationMask = models(oneDesign(0)).RotationMask Then Set template = mSharedTemplate
            End If
        End If
        If template Is Nothing Then
            Set template = BuildMode(models, oneDesign, 1, oneCut, _
                IIf(cutCount = 0, 0, 1), quantities, settings, sourceDocument, layoutMode, planner)
            If cutCount = 1 Then
                Set mSharedTemplate = template
                mTemplateCutIndex = cutIndexes(0)
                mTemplateRotationMask = models(oneDesign(0)).RotationMask
            End If
        End If
        Set result = DistributeTemplate(template, models, designIndexes, designCount, _
            cutIndexes, cutCount)
        If TemplateFits(result, planner, settings, sourceDocument) Then
            AADebugPlan "Quantity balanced template", result
            Set BuildGroup = result
            Exit Function
        End If
        ' Bounds artwork dapat berbeda meskipun Cut Line sama.
        AADebugWrite "Quantity", "template rejected: area or mark; searching balanced quotas"
    End If

    ' Bentuk berbeda: cari quota penuh yang muat dalam SATU area, bukan spillover.
    ' Pencarian memakai heuristic planner yang sama; bukan optimum nesting global.
    For i = 0 To designCount - 1
        quantities(designIndexes(i)) = 0
    Next i
    low = 0: high = AA_MAX_OBJECTS_PER_SOURCE
    If cutCount = 1 Then high = high \ designCount
    Set result = New AALayoutPlan
    result.Configure layoutMode
    middle = (high + 1) \ 2
    If layoutMode = AA_MODE_MEDIUM Then
        Set mQuotaPlanner = New AAMediumPlanner
        mQuotaPlanner.BeginQuotaSearch
        middle = InitialQuota(models, designIndexes, designCount, settings, high)
        expanding = True
        AADebugWrite "Quantity", "initial quota per Design=" & CStr(middle) & _
            "; search limit=" & CStr(high) & "; max difference=" & CStr(AA_QUANTITY_MAX_DIFFERENCE)
    End If
    Do While low < high
        For i = 0 To designCount - 1
            quantities(designIndexes(i)) = middle
        Next i
        expected = middle * designCount
        Set candidate = TryQuota(models, designIndexes, designCount, cutIndexes, cutCount, _
            quantities, expected, settings, sourceDocument, layoutMode, planner)
        If candidate Is Nothing Then
            high = middle - 1
            expanding = False
        Else
            low = middle
            Set result = candidate
        End If
        If expanding Then
            ' Perkiraan bounds bukan batas kapasitas: bila muat, cari di atasnya.
            middle = low * 2
            If middle > high Then middle = high
        Else
            middle = (low + high + 1) \ 2
            ' Hasil parsial hanya memilih probe, bukan menurunkan batas high.
            ' Setelah dua hint, wajib satu langkah biner agar rentang tetap cepat menyusut.
            If layoutMode = AA_MODE_MEDIUM And candidate Is Nothing And low < high Then
                hint = mLastCapacityCount \ designCount
                If hint > low And hint <= high And consecutiveHints < 2 Then
                    middle = hint
                    consecutiveHints = consecutiveHints + 1
                    AADebugWrite "Quantity", "partial hint per Design=" & CStr(hint) & _
                        "; partial=" & CStr(mLastCapacityCount) & "; bracket=" & _
                        CStr(low) & ".." & CStr(high)
                Else
                    consecutiveHints = 0
                End If
            Else
                consecutiveHints = 0
            End If
        End If
    Loop
    For i = 0 To designCount - 1
        quantities(designIndexes(i)) = low
    Next i
    expected = low * designCount
    ' Dua putaran surplus; semua quota tetap dalam [low, low + 2]. Putaran
    ' pertama memberi tiap Design kesempatan sebelum salah satu mendapat +2.
    ReDim failedAtTotal(0 To UBound(models))
    For i = 0 To designCount - 1
        failedAtTotal(designIndexes(i)) = -1
    Next i
    For pass = 1 To AA_QUANTITY_MAX_DIFFERENCE
        For i = 0 To designCount - 1
            index = designIndexes(i)
            If quantities(index) < AA_MAX_OBJECTS_PER_SOURCE And _
                (cutCount <> 1 Or expected < AA_MAX_OBJECTS_PER_SOURCE) And _
                failedAtTotal(index) <> expected Then
                quantities(index) = quantities(index) + 1
                Set candidate = TryQuota(models, designIndexes, designCount, cutIndexes, cutCount, _
                    quantities, expected + 1, settings, sourceDocument, layoutMode, planner)
                If candidate Is Nothing Then
                    quantities(index) = quantities(index) - 1
                    failedAtTotal(index) = expected
                Else
                    expected = expected + 1
                    Set result = candidate
                End If
            End If
        Next i
    Next pass
    If result.DesignCount = 0 Then Err.Raise vbObjectError + 5333, "AAQuantityPlanner", _
        "Tidak ada Design kelompok yang dapat ditempatkan dalam satu container."
    If layoutMode = AA_MODE_MEDIUM Then
        AADebugWrite "Quantity", "optimizing final quotas; target=" & CStr(expected)
        Set candidate = TryQuota(models, designIndexes, designCount, cutIndexes, cutCount, _
            quantities, expected, settings, sourceDocument, layoutMode, planner, False)
        If Not candidate Is Nothing Then
            Set result = candidate
        Else
            ' Rencana kapasitas yang sudah lolos tetap menjadi hasil aman.
            AADebugWrite "Quantity", "final optimization rejected; retained capacity plan"
        End If
    End If
    For i = 0 To designCount - 1
        AADebugWrite "Quantity", "Design=" & CStr(designIndexes(i)) & "; allocated=" & _
            CStr(quantities(designIndexes(i)))
    Next i
    AADebugPlan "Quantity balanced quotas", result
    Set mQuotaPlanner = Nothing
    Set BuildGroup = result
End Function

' Titik awal murah dari footprint + gap, bukan klaim kapasitas maksimum.
' Kontur dapat saling mengisi ruang bbox; loop akan memperbesar quota bila muat.
Private Function InitialQuota(ByRef models() As AAModel, ByRef designIndexes() As Long, _
                              ByVal designCount As Long, ByVal settings As AASettings, _
                              ByVal maximum As Long) As Long
    Dim i As Long, h As Double, v As Double, unitArea As Double, estimate As Double
    Dim model As AAModel
    If maximum <= 0 Then Exit Function
    For i = 0 To designCount - 1
        Set model = models(designIndexes(i))
        model.GetBoundaryGaps AA_MODE_STRAIGHT, h, v
        If h < 0# Then h = 0#
        If v < 0# Then v = 0#
        unitArea = unitArea + (model.FootprintWidth + h) * (model.FootprintHeight + v)
    Next i
    InitialQuota = 1
    If unitArea <= 0# Then Exit Function
    estimate = Fix((settings.AreaWidthMM * settings.AreaHeightMM) / unitArea)
    If estimate >= maximum Then
        InitialQuota = maximum
    ElseIf estimate > 1# Then
        InitialQuota = CLng(estimate)
    End If
End Function

Private Function TemplateFits(ByVal plan As AALayoutPlan, ByVal planner As AAPlanner, _
                              ByVal settings As AASettings, ByVal sourceDocument As Document) As Boolean
    ' CopySlot mempertahankan collision dan posisi template yang sudah lolos
    ' planner dengan gap kontur mm. Hanya bounds artwork pengganti yang perlu
    ' diperiksa kembali; pembagian Design tidak mengubah posisi atau gap.
    TemplateFits = planner.FinalPlanFits(plan, settings, sourceDocument)
End Function

Private Function CanUseTemplate(ByRef models() As AAModel, ByRef designIndexes() As Long, _
                                ByVal designCount As Long, ByVal cutCount As Long) As Boolean
    Dim first As AAModel, other As AAModel, i As Long
    Set first = models(designIndexes(0))
    For i = 1 To designCount - 1
        If first.RotationMask <> models(designIndexes(i)).RotationMask Then Exit Function
    Next i
    If cutCount <> 1 Then
        For i = 1 To designCount - 1
            Set other = models(designIndexes(i))
            If Abs(first.DesignWidth - other.DesignWidth) > 0.0001 Or _
               Abs(first.DesignHeight - other.DesignHeight) > 0.0001 Then Exit Function
            If cutCount = 0 Then
                If Not first.DesignBoundary.SameGeometry(other.DesignBoundary) Then Exit Function
            Else
                If Not first.CanShareCutWith(other) Then Exit Function
            End If
        Next i
    End If
    CanUseTemplate = True
End Function

Private Function DistributeTemplate(ByVal template As AALayoutPlan, ByRef models() As AAModel, _
                                    ByRef designIndexes() As Long, ByVal designCount As Long, _
                                    ByRef cutIndexes() As Long, ByVal cutCount As Long) As AALayoutPlan
    Dim result As AALayoutPlan, source As AAPlacement, placed As AAPlacement
    Dim roleIndex As Long, quota As Long, slot As Long, j As Long, cutIndex As Long
    Set result = New AALayoutPlan
    result.Configure template.LayoutMode
    result.SetRegularPattern template.HasRegularPattern
    slot = 1
    For roleIndex = 0 To designCount - 1
        quota = template.DesignCount \ designCount
        If roleIndex < template.DesignCount Mod designCount Then quota = quota + 1
        cutIndex = -1
        If cutCount = 1 Then
            cutIndex = cutIndexes(0)
        ElseIf cutCount = designCount Then
            cutIndex = cutIndexes(roleIndex)
        End If
        For j = 1 To quota
            Set source = template.Placements(slot)
            Set placed = CopySlot(source, models(designIndexes(roleIndex)), _
                designIndexes(roleIndex), cutIndex)
            result.Add placed
            slot = slot + 1
        Next j
        AADebugWrite "Quantity", "Design=" & CStr(designIndexes(roleIndex)) & "; allocated=" & CStr(quota)
    Next roleIndex
    Set DistributeTemplate = result
End Function

Private Function CopySlot(ByVal source As AAPlacement, ByVal model As AAModel, _
                          ByVal designIndex As Long, ByVal cutIndex As Long) As AAPlacement
    Dim orientation As AARotationVariant, leftX As Double, topY As Double
    Dim result As AAPlacement
    If source.UsesCenter Then
        Set orientation = model.RotationVariant(source.RotationDelta)
        ' Koordinat rencana, bukan bounds kurva hasil pembulatan CorelDRAW.
        leftX = source.CutLeft - orientation.InsetX
        topY = source.CutTop + orientation.InsetY
        Set result = orientation.MakePlacement(designIndex, cutIndex, 0, _
            leftX, topY, source.GapH, source.GapV)
    Else
        leftX = source.DesignLeft: topY = source.DesignTop
        If cutIndex >= 0 Then
            leftX = source.CutLeft - (model.DesignWidth - model.CutWidth) / 2#
            topY = source.CutTop + (model.DesignHeight - model.CutHeight) / 2#
        End If
        Set result = model.MakePlacement(designIndex, cutIndex, 0, leftX, topY)
    End If
    ' CanUseTemplate memastikan siluet acuan sama. Simpan collision asli agar
    ' pergantian Design tidak menggeser slot atau menambah gap antar Design.
    ' Geometry tidak dimutasi; AAPlacement.Translate mengganti referensinya.
    result.SetCollisionFootprint source.CollisionFootprint
    result.SetContourData source.Contour, source.GapH, source.GapV, source.UsesContourGap
    result.SetMarkAreaAllowance source.UsesMarkAreaAllowance
    Set CopySlot = result
End Function

Private Function TryQuota(ByRef models() As AAModel, ByRef designIndexes() As Long, _
                          ByVal designCount As Long, ByRef cutIndexes() As Long, ByVal cutCount As Long, _
                          ByRef quantities() As Long, ByVal expected As Long, ByVal settings As AASettings, _
                          ByVal sourceDocument As Document, ByVal layoutMode As Long, _
                          ByVal planner As AAPlanner, _
                          Optional ByVal capacityOnly As Boolean = True) As AALayoutPlan
    Dim candidate As AALayoutPlan, number As Long, description As String, source As String
    Dim perfStarted As Double, phase As String, capacityTarget As Long
    phase = "full": capacityTarget = -1
    If layoutMode = AA_MODE_MEDIUM And capacityOnly Then
        phase = "capacity": capacityTarget = expected
    End If
    perfStarted = AADebugPerfBegin()
    mLastCapacityCount = 0
    On Error GoTo QuotaFailed
    AAProgressTrial
    Set candidate = BuildMode(models, designIndexes, designCount, cutIndexes, cutCount, _
        quantities, settings, sourceDocument, layoutMode, planner, capacityTarget)
    AADebugStageEnd "quota target=" & CStr(expected) & "; phase=" & phase, perfStarted
    mLastCapacityCount = candidate.DesignCount
    If candidate.MatchesQuotas(designIndexes, designCount, quantities, expected) Then
        Set TryQuota = candidate
        AADebugWrite "Quantity", "quota target=" & CStr(expected) & " -> fits"
    Else
        AADebugWrite "Quantity", "quota target=" & CStr(expected) & " -> capacity"
    End If
    Exit Function
QuotaFailed:
    number = Err.Number: description = Err.Description: source = Err.Source
    AADebugStageEnd "quota target=" & CStr(expected) & "; phase=" & phase & " failed", perfStarted
    Select Case number
        Case vbObjectError + 5210, vbObjectError + 5230, vbObjectError + 5231, _
             vbObjectError + 5233, vbObjectError + 5234, vbObjectError + 5241
            AADebugWrite "Quantity", "quota target=" & CStr(expected) & " -> rejected: " & description
        Case Else
            Err.Raise number, source, description
    End Select
End Function

Private Function BuildMode(ByRef models() As AAModel, ByRef designIndexes() As Long, _
                           ByVal designCount As Long, ByRef cutIndexes() As Long, ByVal cutCount As Long, _
                           ByRef quantities() As Long, ByVal settings As AASettings, _
                           ByVal sourceDocument As Document, ByVal layoutMode As Long, _
                           ByVal planner As AAPlanner, _
                           Optional ByVal capacityTarget As Long = -1) As AALayoutPlan
    If layoutMode = AA_MODE_MEDIUM Then
        If mQuotaPlanner Is Nothing Then
            Set BuildMode = planner.BuildMedium(models, designIndexes, designCount, _
                cutIndexes, cutCount, quantities, settings, sourceDocument, True, capacityTarget)
        Else
            Set BuildMode = mQuotaPlanner.Build(models, designIndexes, designCount, _
                cutIndexes, cutCount, quantities, settings, sourceDocument, planner, True, capacityTarget)
        End If
    Else
        Set BuildMode = planner.BuildMinimum(models, designIndexes, designCount, _
            cutIndexes, cutCount, quantities, settings, sourceDocument, True)
    End If
End Function
