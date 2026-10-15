Option Explicit

Private Const AA_ROW_TOLERANCE_FACTOR As Double = 0.25
Private Const AA_COLUMN_TOLERANCE_FACTOR As Double = 0.25
Private Const AA_CONTACT_TOLERANCE_MM As Double = 0.001
Private Const AA_SAFETY_SCAN_MM As Double = 0.25
Private Const AA_MAX_ALIGNMENT_PROBES As Long = 200000

Private mProbes As Long

Public Sub Align(ByVal batch As Collection)
    Dim referenceLayout As AAOptimizeLayout, rows As Collection, columns As Collection
    Dim members As Collection, prefix As Collection, settled As Collection
    Dim i As Long, j As Long, index As Long, anchor As Long
    Set referenceLayout = batch(1)
    mProbes = 0
    ' Kedua keanggotaan dicatat sebelum posisi rencana pertama berubah.
    Set rows = SnapshotGroups(referenceLayout, True)
    Set columns = SnapshotGroups(referenceLayout, False)
    For i = 1 To rows.Count
        AAProgressStage "Optimize Minimum: baris", CStr(i) & " / " & CStr(rows.Count)
        Set members = rows(i)
        anchor = CLng(members(1))
        Set prefix = New Collection
        prefix.Add anchor
        For j = 2 To members.Count
            index = CLng(members(j))
            PlaceAfter batch, index, CLng(members(j - 1)), True, _
                CenterY(referenceLayout.Placement(anchor)), prefix, prefix
            prefix.Add index
        Next j
    Next i

    Set settled = New Collection
    For i = 1 To columns.Count
        Set members = columns(i)
        settled.Add CLng(members(1))
    Next i
    ' Anchor kolom mengikuti hasil baris. Semua anchor ini tetap di tempat.
    BeginTrial batch, settled
    For i = 1 To settled.Count
        If Not Fits(batch, CLng(settled(i)), 0#, 0#, True) Then NoRoom
    Next i
    For i = 1 To columns.Count
        AAProgressStage "Optimize Minimum: kolom", CStr(i) & " / " & CStr(columns.Count)
        Set members = columns(i)
        anchor = CLng(members(1))
        Set prefix = New Collection
        prefix.Add anchor
        For j = 2 To members.Count
            index = CLng(members(j))
            PlaceAfter batch, index, CLng(members(j - 1)), False, _
                CenterX(referenceLayout.Placement(anchor)), prefix, settled
            prefix.Add index
            settled.Add index
        Next j
    Next i
    AADebugWrite "Optimize Minimum", "rows=" & CStr(rows.Count) & _
        "; columns=" & CStr(columns.Count) & "; probes=" & CStr(mProbes)
End Sub

Private Function SnapshotGroups(ByVal layout As AAOptimizeLayout, ByVal horizontal As Boolean) As Collection
    Dim result As Collection, groups As Collection, members As Collection
    Dim order() As Long, memberIndices() As Long, groupOrder() As Long, used() As Boolean
    Dim count As Long, i As Long, j As Long, k As Long, anchor As Long, memberCount As Long
    Dim p As AAPlacement, tolerance As Double, axis As Double, value As Double
    Set groups = New Collection
    count = layout.Count
    ReDim order(1 To count): ReDim memberIndices(1 To count): ReDim used(1 To count)
    For i = 1 To count
        order(i) = i
    Next i
    ' Baris: calon anchor paling kiri. Kolom: calon anchor paling atas.
    SortMembers layout, order, count, horizontal
    For i = 1 To count
        AAProgressPulse
        anchor = order(i)
        If Not used(anchor) Then
            Set p = layout.OriginalPlacement(anchor)
            If horizontal Then
                axis = CenterY(p): tolerance = p.CollisionFootprint.Height * AA_ROW_TOLERANCE_FACTOR
            Else
                axis = CenterX(p): tolerance = p.CollisionFootprint.Width * AA_COLUMN_TOLERANCE_FACTOR
            End If
            memberCount = 0
            For j = i To count
                k = order(j)
                If Not used(k) Then
                    Set p = layout.OriginalPlacement(k)
                    If horizontal Then value = CenterY(p) Else value = CenterX(p)
                    If Abs(value - axis) <= tolerance + 0.000000001 Then
                        memberCount = memberCount + 1
                        memberIndices(memberCount) = k
                        used(k) = True
                    End If
                End If
            Next j
            Set members = New Collection
            For j = 1 To memberCount
                members.Add memberIndices(j)
            Next j
            groups.Add members
        End If
    Next i
    ReDim groupOrder(1 To groups.Count)
    For i = 1 To groups.Count
        groupOrder(i) = i
    Next i
    For i = 2 To groups.Count
        AAProgressPulse
        k = groupOrder(i): j = i - 1
        Do While j >= 1
            If Not GroupBefore(layout, groups(k), groups(groupOrder(j)), horizontal) Then Exit Do
            groupOrder(j + 1) = groupOrder(j): j = j - 1
        Loop
        groupOrder(j + 1) = k
    Next i
    Set result = New Collection
    For i = 1 To groups.Count
        result.Add groups(groupOrder(i))
    Next i
    Set SnapshotGroups = result
End Function

Private Sub SortMembers(ByVal layout As AAOptimizeLayout, ByRef order() As Long, _
                        ByVal count As Long, ByVal horizontal As Boolean)
    Dim i As Long, j As Long, index As Long
    Dim a As AAPlacement, b As AAPlacement, comesFirst As Boolean
    For i = 2 To count
        AAProgressPulse
        index = order(i): j = i - 1
        Set a = layout.OriginalPlacement(index)
        Do While j >= 1
            Set b = layout.OriginalPlacement(order(j))
            If horizontal Then
                comesFirst = (CenterX(a) < CenterX(b))
                If CenterX(a) = CenterX(b) Then comesFirst = (CenterY(a) > CenterY(b))
            Else
                comesFirst = (CenterY(a) > CenterY(b))
                If CenterY(a) = CenterY(b) Then comesFirst = (CenterX(a) < CenterX(b))
            End If
            If Not comesFirst Then Exit Do
            order(j + 1) = order(j): j = j - 1
        Loop
        order(j + 1) = index
    Next i
End Sub

Private Function GroupBefore(ByVal layout As AAOptimizeLayout, ByVal a As Collection, _
                             ByVal b As Collection, ByVal horizontal As Boolean) As Boolean
    Dim p As AAPlacement, i As Long, first As Double, second As Double
    If horizontal Then
        first = -1E+30: second = -1E+30
        For i = 1 To a.Count
            Set p = layout.OriginalPlacement(CLng(a(i)))
            If CenterY(p) > first Then first = CenterY(p)
        Next i
        For i = 1 To b.Count
            Set p = layout.OriginalPlacement(CLng(b(i)))
            If CenterY(p) > second Then second = CenterY(p)
        Next i
        GroupBefore = (first > second)
    Else
        GroupBefore = (CenterX(layout.OriginalPlacement(CLng(a(1)))) < _
                       CenterX(layout.OriginalPlacement(CLng(b(1)))))
    End If
End Function

Private Sub PlaceAfter(ByVal batch As Collection, ByVal index As Long, ByVal previousIndex As Long, _
                       ByVal horizontal As Boolean, ByVal aligned As Double, _
                       ByVal prefix As Collection, ByVal settled As Collection)
    Dim referenceLayout As AAOptimizeLayout, layout As AAOptimizeLayout
    Dim p As AAPlacement, other As AAPlacement, origin As Double, limit As Double
    Dim safe As Double, blocked As Double, trial As Double, stepMM As Double
    Dim dx As Double, dy As Double, gap As Double, needed As Double, nextDistance As Double
    Dim i As Long, distance As Double, candidate As Double
    Set referenceLayout = batch(1)
    Set other = referenceLayout.Placement(previousIndex)
    If horizontal Then origin = CenterX(other) Else origin = CenterY(other)
    limit = 1E+30
    For Each layout In batch
        Set p = layout.Placement(index)
        If horizontal Then
            needed = layout.AreaWidth - p.Footprint.RightX + CenterX(p) - origin
        Else
            needed = p.Footprint.BottomY + origin - CenterY(p)
        End If
        If needed < limit Then limit = needed
    Next layout
    If limit < AA_CONTACT_TOLERANCE_MM Then NoRoom
    Set p = referenceLayout.Placement(index)
    For i = 1 To prefix.Count
        Set other = referenceLayout.Placement(CLng(prefix(i)))
        If horizontal Then
            gap = p.GapH: If other.GapH > gap Then gap = other.GapH
            needed = other.CollisionFootprint.RightX + gap + p.CollisionFootprint.Width / 2# - origin
        Else
            gap = p.GapV: If other.GapV > gap Then gap = other.GapV
            needed = origin - other.CollisionFootprint.BottomY + gap + p.CollisionFootprint.Height / 2#
        End If
        If needed > safe Then safe = needed
    Next i
    safe = safe + AA_CONTACT_TOLERANCE_MM
    If safe > limit Then safe = limit
    BeginTrial batch, prefix
    Offsets p, horizontal, origin, aligned, safe, dx, dy
    If Not Fits(batch, index, dx, dy, False) Then NoRoom
    stepMM = p.CollisionFootprint.Height / 16#
    If horizontal Then stepMM = p.CollisionFootprint.Width / 16#
    If stepMM < 1# Then stepMM = 1#
    Do
        trial = safe - stepMM
        If trial < AA_CONTACT_TOLERANCE_MM Then trial = AA_CONTACT_TOLERANCE_MM
        Offsets p, horizontal, origin, aligned, trial, dx, dy
        If Fits(batch, index, dx, dy, False) Then
            safe = trial
            If trial = AA_CONTACT_TOLERANCE_MM Then Exit Do
        Else
            blocked = trial
            Do While safe - blocked > AA_CONTACT_TOLERANCE_MM
                trial = (safe + blocked) / 2#
                Offsets p, horizontal, origin, aligned, trial, dx, dy
                If Fits(batch, index, dx, dy, False) Then safe = trial Else blocked = trial
            Loop
            Exit Do
        End If
    Loop
    ' Setelah gap pada garis terpenuhi, periksa area/mark dan posisi yang
    ' sudah selesai. Titik pusat tetangga ikut dicoba agar gap H/V berbeda
    ' tetap dapat menemukan alignment tepat pada baris yang sudah selesai.
    BeginTrial batch, settled
    distance = safe
    Do
        Offsets p, horizontal, origin, aligned, distance, dx, dy
        If Fits(batch, index, dx, dy, True) Then Exit Do
        If distance >= limit Then NoRoom
        nextDistance = distance + AA_SAFETY_SCAN_MM
        If nextDistance > limit Then nextDistance = limit
        For i = 1 To settled.Count
            Set other = referenceLayout.Placement(CLng(settled(i)))
            If horizontal Then candidate = CenterX(other) - origin Else candidate = origin - CenterY(other)
            If candidate > distance + 0.000000001 And candidate < nextDistance Then nextDistance = candidate
        Next i
        distance = nextDistance
    Loop
    For Each layout In batch
        layout.MovePlan index, dx, dy
    Next layout
End Sub

Private Sub Offsets(ByVal p As AAPlacement, ByVal horizontal As Boolean, _
                    ByVal origin As Double, ByVal aligned As Double, ByVal distance As Double, _
                    ByRef dx As Double, ByRef dy As Double)
    If horizontal Then
        dx = origin + distance - CenterX(p): dy = aligned - CenterY(p)
    Else
        dx = aligned - CenterX(p): dy = origin - distance - CenterY(p)
    End If
End Sub

Private Sub BeginTrial(ByVal batch As Collection, ByVal settled As Collection)
    Dim layout As AAOptimizeLayout
    For Each layout In batch
        layout.BeginMinimumTrial settled
    Next layout
End Sub

Private Function Fits(ByVal batch As Collection, ByVal index As Long, _
                      ByVal dx As Double, ByVal dy As Double, ByVal safety As Boolean) As Boolean
    Dim layout As AAOptimizeLayout
    AAProgressPulse
    mProbes = mProbes + 1
    If mProbes > AA_MAX_ALIGNMENT_PROBES Then Err.Raise vbObjectError + 5373, _
        "AAOptimizeMinimum", "Pencarian Optimize Minimum mencapai batas. Coba mode Medium atau rapikan susunan awal."
    For Each layout In batch
        If Not layout.MinimumCanMove(index, dx, dy, safety) Then Exit Function
    Next layout
    Fits = True
End Function

Private Function CenterX(ByVal p As AAPlacement) As Double
    CenterX = (p.CollisionFootprint.LeftX + p.CollisionFootprint.RightX) / 2#
End Function

Private Function CenterY(ByVal p As AAPlacement) As Double
    CenterY = (p.CollisionFootprint.BottomY + p.CollisionFootprint.TopY) / 2#
End Function

Private Sub NoRoom()
    Err.Raise vbObjectError + 5372, "AAOptimizeMinimum", _
        "Optimize Minimum tidak menemukan posisi yang memenuhi Gap, Area, atau zona mark. Susunan awal dipertahankan."
End Sub
