Option Explicit

Private Const AA_MAX_OBJECTS_PER_SOURCE As Long = 5000
Private mGroups As Collection
Private mLabels As Collection
Private mOrigins() As Long
Private mRemaining() As Long
Private mUsed() As Long
Private mCutForDesign() As Long
Private mOrder() As Long
Private mOrderCount As Long
Private mSealed() As Boolean
Private mPullFree() As Boolean
Private mAllocator As AAQuantityPlanner

' Hanya /n dan Quantity bermodifier. Entry biasa tetap melalui jalur template lama.
Public Function Build(ByVal objects As AAObjectStore, ByRef models() As AAModel, _
                       ByRef designs() As Long, ByVal designCount As Long, _
                       ByRef cuts() As Long, ByVal cutCount As Long, _
                       ByVal settings As AASettings, ByVal doc As Document, _
                       ByVal layoutMode As Long, ByVal planner As AAPlanner) As AALayoutPlan
    Dim result As AALayoutPlan, part As AALayoutPlan, members As Collection
    Dim selected() As Long, selectedCuts() As Long, requested() As Long, limits() As Long
    Dim count As Long, i As Long, index As Long, group As Long
    Dim before As Double, activeOrigins() As Boolean
    ReDim mOrigins(0 To objects.Count - 1): ReDim mRemaining(0 To objects.Count - 1)
    ReDim mUsed(0 To objects.Count - 1): ReDim mCutForDesign(0 To objects.Count - 1)
    ReDim mOrder(0 To designCount - 1): mOrderCount = 0
    For i = 0 To designCount - 1
        index = designs(i): mCutForDesign(index) = -1
        If cutCount = 1 Then mCutForDesign(index) = cuts(0)
        If cutCount > 1 Then mCutForDesign(index) = cuts(i)
        If objects.QuantityAt(index) > AA_MAX_OBJECTS_PER_SOURCE Then Err.Raise _
            vbObjectError + 5210, "AAQuantityFlow", "Quantity satu Design melebihi batas 5000 objek."
        If objects.QuantityAt(index) > 0 Then mRemaining(index) = objects.QuantityAt(index)
    Next i
    BuildGroups objects, designs, designCount
    Set mAllocator = New AAQuantityPlanner
    Set result = New AALayoutPlan: result.Configure layoutMode: result.SetRegularPattern True
    For group = 1 To mGroups.Count
        AAProgressGroup CStr(mLabels(group))
        AAProgressPulse
        ResetRequests objects.Count, selected, selectedCuts, requested, limits, count
        Set members = mGroups(group)
        ' Milik tujuan lebih dulu; carry tidak menggantikan target tetap tujuan.
        For i = 1 To members.Count
            index = CLng(members(i))
            If mRemaining(index) > 0 Then Queue index, mRemaining(index), selected, requested, limits, count
        Next i
        If Not mSealed(group) Then
            For i = 0 To mOrderCount - 1
                index = mOrder(i)
                If mOrigins(index) < group And mRemaining(index) > 0 Then _
                    Queue index, mRemaining(index), selected, requested, limits, count
            Next i
        End If
        For i = 1 To members.Count
            index = CLng(members(i))
            If objects.QuantityAt(index) < 0 And mUsed(index) < AA_MAX_OBJECTS_PER_SOURCE Then _
                Queue index, -1, selected, requested, limits, count
        Next i
        Set part = Allocate(models, selected, selectedCuts, count, requested, limits, _
            settings, doc, layoutMode, planner)
        If mPullFree(group) Then Set part = PullFutureFree(part, group, objects, models, designs, _
            designCount, settings, doc, layoutMode, planner)
        CommitPart part, result, objects
    Next group
    ' Filler tidak membuka container sendiri setelah target tetap habis.
    Do While PendingCount() > 0#
        AAProgressGroup "Sisa Quantity - container tambahan"
        AAProgressPulse
        before = PendingCount()
        ReDim activeOrigins(1 To mGroups.Count)
        ResetRequests objects.Count, selected, selectedCuts, requested, limits, count
        For i = 0 To mOrderCount - 1
            index = mOrder(i)
            If mRemaining(index) > 0 Then
                activeOrigins(mOrigins(index)) = True
                Queue index, mRemaining(index), selected, requested, limits, count
            End If
        Next i
        ' Filler anggota kelompok tetap boleh mengikuti sisa target kelompoknya.
        For i = 0 To mOrderCount - 1
            index = mOrder(i)
            If activeOrigins(mOrigins(index)) And objects.QuantityAt(index) < 0 And _
               mUsed(index) < AA_MAX_OBJECTS_PER_SOURCE Then Queue index, -1, selected, requested, limits, count
        Next i
        Set part = Allocate(models, selected, selectedCuts, count, requested, limits, _
            settings, doc, layoutMode, planner)
        CommitPart part, result, objects
        If PendingCount() >= before Then Err.Raise vbObjectError + 5338, "AAQuantityFlow", _
            "Sisa Quantity tidak muat dalam area kosong. Periksa ukuran, gap dan mark: " & FirstPending(objects)
    Loop
    If Not planner.FinalPlanFits(result, settings, doc) Then Err.Raise vbObjectError + 5331, _
        "AAQuantityFlow", "Susunan Quantity tidak memenuhi batas area atau mark."
    Set Build = result
End Function

Private Sub BuildGroups(ByVal objects As AAObjectStore, ByRef designs() As Long, ByVal designCount As Long)
    Dim numbers As Object, keys As Variant, ordered() As Long, members As Collection
    Dim i As Long, j As Long, value As Long, index As Long, group As Long
    Set numbers = CreateObject("Scripting.Dictionary")
    Set mGroups = New Collection: Set mLabels = New Collection
    For i = 0 To designCount - 1
        value = objects.QuantityGroupAt(designs(i))
        If value > 0 Then numbers(CStr(value)) = value
    Next i
    If numbers.Count > 0 Then
        keys = numbers.Keys: ReDim ordered(0 To numbers.Count - 1)
        For i = 0 To numbers.Count - 1
            ordered(i) = CLng(keys(i))
        Next i
        For i = 1 To UBound(ordered)
            AAProgressPulse
            j = i
            Do While j > 0
                If ordered(j - 1) <= ordered(j) Then Exit Do
                value = ordered(j - 1): ordered(j - 1) = ordered(j): ordered(j) = value
                j = j - 1
            Loop
        Next i
        For i = 0 To UBound(ordered)
            AAProgressPulse
            Set members = New Collection
            For j = 0 To designCount - 1
                index = designs(j)
                If objects.QuantityGroupAt(index) = ordered(i) Then members.Add index
            Next j
            mGroups.Add members: mLabels.Add "Kelompok /" & CStr(ordered(i))
        Next i
    End If
    ' Angka bermodifier tanpa /n memiliki container sendiri, menurut ListBox.
    For i = 0 To designCount - 1
        index = designs(i)
        If objects.QuantityGroupAt(index) = 0 And objects.HasExtendedQuantityAt(index) Then
            Set members = New Collection: members.Add index
            mGroups.Add members: mLabels.Add "Design /" & CStr(objects.StaticIDAt(index))
        End If
    Next i
    ReDim mSealed(1 To mGroups.Count): ReDim mPullFree(1 To mGroups.Count)
    For group = 1 To mGroups.Count
        Set members = mGroups(group)
        For i = 1 To members.Count
            index = CLng(members(i)): mOrigins(index) = group
            mOrder(mOrderCount) = index: mOrderCount = mOrderCount + 1
            If objects.QuantityAt(index) <> 0 Then
                mSealed(group) = mSealed(group) Or objects.QuantitySealedAt(index)
                mPullFree(group) = mPullFree(group) Or objects.QuantityPullFreeAt(index)
            End If
        Next i
        If mSealed(group) And mPullFree(group) Then Err.Raise vbObjectError + 5339, _
            "AAQuantityFlow", CStr(mLabels(group)) & " memakai * dan & bersamaan. Pilih satu aturan."
    Next group
End Sub

Private Sub ResetRequests(ByVal entryCount As Long, ByRef selected() As Long, _
                           ByRef cuts() As Long, ByRef requested() As Long, _
                           ByRef limits() As Long, ByRef count As Long)
    ReDim selected(0 To entryCount - 1): ReDim cuts(0 To entryCount - 1)
    ReDim requested(0 To entryCount - 1): ReDim limits(0 To entryCount - 1)
    count = 0
End Sub

Private Sub Queue(ByVal index As Long, ByVal amount As Long, ByRef selected() As Long, _
                   ByRef requested() As Long, ByRef limits() As Long, ByRef count As Long)
    selected(count) = index: count = count + 1
    requested(index) = amount: limits(index) = AA_MAX_OBJECTS_PER_SOURCE - mUsed(index)
End Sub

Private Function Allocate(ByRef models() As AAModel, ByRef selected() As Long, _
                           ByRef cuts() As Long, ByVal count As Long, ByRef requested() As Long, _
                           ByRef limits() As Long, ByVal settings As AASettings, ByVal doc As Document, _
                           ByVal layoutMode As Long, ByVal planner As AAPlanner) As AALayoutPlan
    Dim result As AALayoutPlan, i As Long, cutCount As Long, sameCut As Boolean
    If count = 0 Then
        Set result = New AALayoutPlan: result.Configure layoutMode
    Else
        sameCut = True
        For i = 0 To count - 1
            cuts(i) = mCutForDesign(selected(i))
            If cuts(i) <> cuts(0) Then sameCut = False
        Next i
        cutCount = count
        If sameCut Then
            cutCount = 1
            If cuts(0) < 0 Then cutCount = 0
        End If
        Set result = mAllocator.BuildAllocatedGroup(models, selected, count, cuts, cutCount, _
            requested, limits, settings, doc, layoutMode, planner)
    End If
    Set Allocate = result
End Function

Private Function PullFutureFree(ByVal current As AALayoutPlan, ByVal group As Long, _
                                 ByVal objects As AAObjectStore, ByRef models() As AAModel, _
                                 ByRef designs() As Long, ByVal designCount As Long, _
                                 ByVal settings As AASettings, ByVal doc As Document, _
                                 ByVal layoutMode As Long, ByVal planner As AAPlanner) As AALayoutPlan
    Dim selected() As Long, cuts() As Long, requested() As Long, limits() As Long
    Dim pinned() As Long, placed As AAPlacement, candidate As AALayoutPlan
    Dim count As Long, i As Long, index As Long, freeCount As Long, valid As Boolean
    Set PullFutureFree = current
    ReDim pinned(0 To objects.Count - 1)
    For Each placed In current.Placements
        pinned(placed.DesignIndex) = pinned(placed.DesignIndex) + 1
    Next placed
    ResetRequests objects.Count, selected, cuts, requested, limits, count
    ' Pertahankan urutan sumber rencana yang sudah terbentuk, termasuk carry.
    For Each placed In current.Placements
        index = placed.DesignIndex
        If requested(index) = 0 Then Queue index, pinned(index), selected, requested, limits, count
    Next placed
    For i = 0 To designCount - 1
        index = designs(i)
        If (mOrigins(index) > group Or mOrigins(index) = 0) And objects.QuantityAt(index) < 0 And _
           mUsed(index) < AA_MAX_OBJECTS_PER_SOURCE Then
            Queue index, -1, selected, requested, limits, count
            freeCount = freeCount + 1
        End If
    Next i
    If freeCount = 0 Then Exit Function
    Set candidate = Allocate(models, selected, cuts, count, requested, limits, settings, doc, layoutMode, planner)
    For Each placed In candidate.Placements
        pinned(placed.DesignIndex) = pinned(placed.DesignIndex) - 1
    Next placed
    valid = True
    For i = 0 To UBound(pinned)
        If pinned(i) > 0 Then valid = False
    Next i
    If valid Then Set PullFutureFree = candidate
End Function

Private Sub CommitPart(ByVal part As AALayoutPlan, ByVal result As AALayoutPlan, ByVal objects As AAObjectStore)
    Dim counts() As Long, placed As AAPlacement, i As Long, areaBase As Long
    If part.AreaCount > 1 Then Err.Raise vbObjectError + 5334, "AAQuantityFlow", _
        "Alokasi Quantity harus menghasilkan satu container per langkah."
    If part.DesignCount = 0 Then Exit Sub
    ReDim counts(0 To objects.Count - 1)
    For Each placed In part.Placements
        counts(placed.DesignIndex) = counts(placed.DesignIndex) + 1
    Next placed
    For i = 0 To UBound(counts)
        If objects.QuantityAt(i) >= 0 Then
            If counts(i) > mRemaining(i) Then Err.Raise vbObjectError + 5340, "AAQuantityFlow", _
                "Alokasi melebihi Quantity Design /" & CStr(objects.StaticIDAt(i))
            mRemaining(i) = mRemaining(i) - counts(i)
        End If
        mUsed(i) = mUsed(i) + counts(i)
        If mUsed(i) > AA_MAX_OBJECTS_PER_SOURCE Then Err.Raise vbObjectError + 5210, _
            "AAQuantityFlow", "Satu Design melebihi batas 5000 objek."
        If counts(i) > 0 Then AADebugWrite "Quantity", "Design /" & CStr(objects.StaticIDAt(i)) & _
            "; allocated=" & CStr(counts(i)) & "; remaining=" & CStr(mRemaining(i))
    Next i
    areaBase = result.AreaCount
    result.SetRegularPattern result.HasRegularPattern And part.HasRegularPattern
    For Each placed In part.Placements
        placed.SetAreaIndex areaBase
        result.Add placed
    Next placed
End Sub

Private Function PendingCount() As Double
    Dim i As Long
    For i = 0 To mOrderCount - 1
        PendingCount = PendingCount + CDbl(mRemaining(mOrder(i)))
    Next i
End Function

Private Function FirstPending(ByVal objects As AAObjectStore) As String
    Dim i As Long, index As Long
    For i = 0 To mOrderCount - 1
        index = mOrder(i)
        If mRemaining(index) > 0 Then
            FirstPending = "Design /" & CStr(objects.StaticIDAt(index)) & "; sisa=" & CStr(mRemaining(index))
            Exit Function
        End If
    Next i
End Function
