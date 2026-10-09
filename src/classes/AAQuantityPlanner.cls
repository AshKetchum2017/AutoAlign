Option Explicit

Private Const AA_MAX_OBJECTS_PER_SOURCE As Long = 5000
Private Const AA_QUANTITY_MAX_DIFFERENCE As Long = 2
Private mSharedTemplate As AALayoutPlan
Private mTemplateCutIndex As Long
Private mQuotaPlanner As AAMediumPlanner
Private mLastCapacityCount As Long

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
                If mTemplateCutIndex = cutIndexes(0) Then Set template = mSharedTemplate
            End If
        End If
        If template Is Nothing Then
            Set template = BuildMode(models, oneDesign, 1, oneCut, _
                IIf(cutCount = 0, 0, 1), quantities, settings, sourceDocument, layoutMode, planner)
            If cutCount = 1 Then
                Set mSharedTemplate = template
                mTemplateCutIndex = cutIndexes(0)
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
    ' planner. Jangan mengubah kontrak gap template, termasuk ZigZag 88% bawaan
    ' Minimum. Hanya bounds artwork pengganti yang perlu diperiksa kembali.
    TemplateFits = planner.FinalPlanFits(plan, settings, sourceDocument)
End Function

Private Function CanUseTemplate(ByRef models() As AAModel, ByRef designIndexes() As Long, _
                                ByVal designCount As Long, ByVal cutCount As Long) As Boolean
    Dim first As AAModel, other As AAModel, i As Long
    Set first = models(designIndexes(0))
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
