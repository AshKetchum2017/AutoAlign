Option Explicit

Private Const AA_OPTIMIZE_PASSES As Long = 3
Private Const AA_OPTIMIZE_STEP_MM As Double = 1#
Private Const AA_OPTIMIZE_STOP_MM As Double = 0.001
Private Const AA_OPTIMIZE_MAX_PROBES As Long = 20000
Private Const AA_OPTIMIZE_LINE_TOLERANCE_MM As Double = 0.001

Public Sub Run(ByVal objects As AAObjectStore, ByVal modelMode As Long, _
               ByVal areaWidthText As String, ByVal areaHeightText As String, _
               ByVal horizontalGapText As String, ByVal verticalGapText As String, _
               ByRef completed As Boolean, ByRef feedback As String, ByRef feedbackStyle As VbMsgBoxStyle)
    Dim settings As AASettings, planner As AAPlanner, forbidden As Collection
    Dim doc As Document, page As Page, selection As ShapeRange, root As Shape, cut As Shape
    Dim designs As New Collection, cuts As New Collection, batches As New Collection
    Dim allLayouts As New Collection, batch As Collection, layout As AAOptimizeLayout, previous As AAOptimizeLayout
    Dim i As Long, j As Long, previousUnit As cdrUnit, previousH As cdrUnit, previousV As cdrUnit
    Dim unitsSaved As Boolean, groupOpen As Boolean, mutationStarted As Boolean, rollbackOK As Boolean
    Dim errorNumber As Long, errorSource As String, errorDescription As String, started As Double
    Dim cleanupError As String

    completed = False: feedback = vbNullString: feedbackStyle = vbExclamation
    On Error GoTo Failed
    AADebugOptimizeStart modelMode
    started = AADebugPerfBegin()
    If objects.Count = 0 Then Err.Raise vbObjectError + 5361, "AAOptimizeEngine", _
        "Daftarkan group susunan ke ListBox terlebih dahulu."
    Set settings = New AASettings
    settings.Configure modelMode, areaWidthText, areaHeightText, horizontalGapText, verticalGapText
    AADebugSettings settings
    If Application.Documents.Count = 0 Then Err.Raise vbObjectError + 5362, "AAOptimizeEngine", _
        "Tidak ada dokumen aktif."
    Set doc = ActiveDocument: Set page = doc.ActivePage
    If Not (doc Is objects.SourceDocument) Or page.Index <> objects.SourcePageIndex Then _
        Err.Raise vbObjectError + 5363, "AAOptimizeEngine", "Aktifkan dokumen dan halaman sumber."
    Set selection = ActiveSelectionRange
    For i = 0 To objects.Count - 1
        Set root = FindRoot(page, objects.StaticIDAt(i))
        If root Is Nothing Then Err.Raise vbObjectError + 5364, "AAOptimizeEngine", _
            "Container terdaftar tidak ditemukan. Perbarui daftar."
        CheckEditable root
        If objects.RoleAt(i) = AA_ROLE_DESIGN Then designs.Add root Else cuts.Add root
    Next i
    If designs.Count > 0 And cuts.Count > 1 And cuts.Count <> designs.Count Then _
        Err.Raise vbObjectError + 5365, "AAOptimizeEngine", _
            "Jumlah container Cut Line harus satu atau sama dengan container Design."
    ' Cut Line saja juga dapat dioptimasi sebagai susunan standalone.
    If designs.Count = 0 Then
        Set designs = cuts
        Set cuts = New Collection
    End If

    previousUnit = doc.Unit: previousH = doc.Rulers.HUnits: previousV = doc.Rulers.VUnits
    unitsSaved = True
    doc.BeginCommandGroup "Auto Align Optimize": groupOpen = True
    doc.Unit = cdrMillimeter: doc.Rulers.HUnits = cdrMillimeter: doc.Rulers.VUnits = cdrMillimeter
    Set planner = New AAPlanner
    Set forbidden = planner.ForbiddenFor(settings, doc)
    For i = 1 To designs.Count
        If i = 1 Or cuts.Count <> 1 Then
            Set batch = New Collection
            batches.Add batch
        End If
        Set cut = Nothing
        If cuts.Count = 1 Then
            Set cut = cuts(1)
        ElseIf cuts.Count > 1 Then
            Set cut = cuts(i)
        End If
        Set layout = New AAOptimizeLayout
        layout.Prepare designs(i), cut, settings, forbidden
        For Each previous In batch
            If Not layout.MatchesTemplate(previous) Then Err.Raise vbObjectError + 5370, _
                "AAOptimizeEngine", "Susunan container Design tidak identik dalam toleransi template 1 mm."
        Next previous
        batch.Add layout: allLayouts.Add layout
        If Not layout.Valid Then Err.Raise vbObjectError + 5366, "AAOptimizeEngine", _
            "Susunan awal tidak memenuhi Area, Gap, atau zona mark. Tidak ada objek yang digeser."
    Next i
    AADebugStageEnd "Optimize prepare", started
    started = AADebugPerfBegin()
    For i = 1 To batches.Count
        Set batch = batches(i)
        AAProgressStage "Merapatkan susunan", "Container " & CStr(i) & " / " & CStr(batches.Count)
        Compact batch
    Next i
    AADebugStageEnd "Optimize compact", started
    AAProgressStage "Memvalidasi susunan"
    For Each layout In allLayouts
        If Not layout.Valid Then Err.Raise vbObjectError + 5367, "AAOptimizeEngine", _
            "Rencana Optimize gagal validasi. Tidak ada objek yang digeser."
    Next layout
    AAProgressStage "Menerapkan pergeseran"
    mutationStarted = True
    For i = 1 To batches.Count
        Set batch = batches(i)
        For j = 1 To batch.Count
            Set layout = batch(j)
            ' Template bersama hanya dipindahkan sekali.
            layout.Apply (j = 1)
        Next j
    Next i
    AAProgressStage "Memeriksa hasil aktual"
    For Each layout In allLayouts
        If Not layout.AppliedValid Then Err.Raise vbObjectError + 5369, "AAOptimizeEngine", _
            "Hasil pergeseran tidak lolos validasi aktual. Posisi semula dipulihkan."
    Next layout
    completed = True
    GoTo Cleanup

Failed:
    errorNumber = Err.Number: errorSource = Err.Source: errorDescription = Err.Description
    AADebugWrite "Optimize", "FAILED: " & CStr(errorNumber) & " " & errorSource & " " & errorDescription
    If mutationStarted Then
        rollbackOK = True
        On Error Resume Next
        For Each layout In allLayouts
            Err.Clear
            layout.Restore
            If Err.Number <> 0 Then rollbackOK = False
        Next layout
        On Error GoTo 0
        If Not rollbackOK Then errorDescription = errorDescription & _
            vbCrLf & "Pemulihan posisi belum lengkap. Gunakan Undo satu kali."
    End If
    feedback = errorDescription
Cleanup:
    On Error Resume Next
    Err.Clear
    If Not selection Is Nothing Then selection.CreateSelection
    If Err.Number <> 0 Then cleanupError = "seleksi"
    Err.Clear
    If unitsSaved Then
        doc.Unit = previousUnit: doc.Rulers.HUnits = previousH: doc.Rulers.VUnits = previousV
    End If
    If Err.Number <> 0 Then cleanupError = cleanupError & " satuan"
    Err.Clear
    If groupOpen Then doc.EndCommandGroup
    If Err.Number <> 0 Then cleanupError = cleanupError & " CommandGroup"
    On Error GoTo 0
    If Len(cleanupError) > 0 Then
        completed = False
        If Len(feedback) > 0 Then feedback = feedback & vbCrLf
        feedback = feedback & "State dokumen belum pulih sepenuhnya:" & cleanupError & "."
        AADebugWrite "Optimize", feedback
    End If
    If completed Then
        AADebugWrite "Optimize", "OK; count and rotations preserved"
        AADebugPerfFinish "Optimize OK"
    Else
        AADebugPerfFinish "Optimize FAILED"
    End If
End Sub

Private Sub Compact(ByVal batch As Collection)
    Dim referenceLayout As AAOptimizeLayout, layout As AAOptimizeLayout
    Dim units As Collection, indices As Collection, p As AAPlacement
    Dim pass As Long, visit As Long, unitIndex As Long, i As Long, direction As Long, moved As Long, probes As Long
    Dim dx As Double, dy As Double, bestX As Double, bestY As Double
    Dim oldArea As Double, oldSpread As Double
    Dim leftX As Double, topY As Double, rightX As Double, bottomY As Double
    Dim unitLeft As Double, unitBottom As Double, unitRight As Double, unitTop As Double
    Dim maximum As Double, safe As Double, blocked As Double, distance As Double, sx As Double, sy As Double
    Set referenceLayout = batch(1)
    Set units = MovementUnits(referenceLayout)
    For pass = 1 To AA_OPTIMIZE_PASSES
        moved = 0
        For visit = 1 To units.Count
            AAProgressPulse
            If probes >= AA_OPTIMIZE_MAX_PROBES Then Exit For
            unitIndex = visit
            If pass Mod 2 = 0 Then unitIndex = units.Count - visit + 1
            Set indices = units(unitIndex)
            For Each layout In batch
                layout.BeginGroupTrial indices
            Next layout
            For direction = 1 To 4
                Stats referenceLayout, Nothing, 0#, 0#, leftX, bottomY, rightX, topY, oldArea, oldSpread
                For i = 1 To indices.Count
                    Set p = referenceLayout.Placement(CLng(indices(i)))
                    If i = 1 Then
                        unitLeft = p.Footprint.LeftX: unitBottom = p.Footprint.BottomY
                        unitRight = p.Footprint.RightX: unitTop = p.Footprint.TopY
                    Else
                        unitLeft = MinValue(unitLeft, p.Footprint.LeftX)
                        unitBottom = MinValue(unitBottom, p.Footprint.BottomY)
                        unitRight = MaxValue(unitRight, p.Footprint.RightX)
                        unitTop = MaxValue(unitTop, p.Footprint.TopY)
                    End If
                Next i
                sx = 0#: sy = 0#
                Select Case direction
                    Case 1
                        sx = -1#: maximum = unitLeft - leftX
                    Case 2
                        sy = 1#: maximum = topY - unitTop
                    Case 3
                        sx = 1#: maximum = rightX - unitRight
                    Case 4
                        sy = -1#: maximum = unitBottom - bottomY
                End Select
                If maximum > AA_OPTIMIZE_STOP_MM And probes < AA_OPTIMIZE_MAX_PROBES Then
                    safe = 0#: blocked = 0#: bestX = 0#: bestY = 0#
                    distance = MinValue(AA_OPTIMIZE_STEP_MM, maximum)
                    Do
                        AAProgressPulse
                        probes = probes + 1
                        dx = sx * distance: dy = sy * distance
                        If Not BatchFits(batch, indices, dx, dy) Then
                            blocked = distance
                            Exit Do
                        End If
                        safe = distance
                        Consider referenceLayout, indices, dx, dy, oldArea, oldSpread, bestX, bestY
                        If safe >= maximum Or probes >= AA_OPTIMIZE_MAX_PROBES Then Exit Do
                        distance = MinValue(safe + AA_OPTIMIZE_STEP_MM, maximum)
                    Loop
                    Do While blocked - safe > AA_OPTIMIZE_STOP_MM And probes < AA_OPTIMIZE_MAX_PROBES
                        distance = (safe + blocked) / 2#
                        dx = sx * distance: dy = sy * distance
                        probes = probes + 1
                        If BatchFits(batch, indices, dx, dy) Then
                            safe = distance
                            Consider referenceLayout, indices, dx, dy, oldArea, oldSpread, bestX, bestY
                        Else
                            blocked = distance
                        End If
                    Loop
                    If bestX <> 0# Or bestY <> 0# Then
                        For Each layout In batch
                            For i = 1 To indices.Count
                                layout.MovePlan CLng(indices(i)), bestX, bestY
                            Next i
                        Next layout
                        moved = moved + indices.Count
                    End If
                End If
            Next direction
        Next visit
        AADebugWrite "Optimize", "pass=" & CStr(pass) & "; moves=" & CStr(moved) & "; probes=" & CStr(probes)
        If moved = 0 Or probes >= AA_OPTIMIZE_MAX_PROBES Then Exit For
    Next pass
    CenterPlan batch
End Sub

Private Sub Consider(ByVal referenceLayout As AAOptimizeLayout, ByVal indices As Collection, _
                     ByVal dx As Double, ByVal dy As Double, ByRef bestArea As Double, _
                     ByRef bestSpread As Double, ByRef bestX As Double, ByRef bestY As Double)
    Dim l As Double, b As Double, r As Double, t As Double, area As Double, spread As Double
    Stats referenceLayout, indices, dx, dy, l, b, r, t, area, spread
    If area < bestArea - 0.000001 Or _
       (Abs(area - bestArea) <= 0.000001 And spread < bestSpread - 0.000001) Then
        bestArea = area: bestSpread = spread: bestX = dx: bestY = dy
    End If
End Sub

Private Function BatchFits(ByVal batch As Collection, ByVal indices As Collection, _
                           ByVal dx As Double, ByVal dy As Double) As Boolean
    Dim layout As AAOptimizeLayout, i As Long
    For Each layout In batch
        For i = 1 To indices.Count
            If Not layout.CanMove(CLng(indices(i)), dx, dy) Then Exit Function
        Next i
    Next layout
    BatchFits = True
End Function

Private Sub CenterPlan(ByVal batch As Collection)
    Dim layout As AAOptimizeLayout, referenceLayout As AAOptimizeLayout
    Dim l As Double, b As Double, r As Double, t As Double, area As Double, spread As Double
    Dim dx As Double, dy As Double, i As Long, valid As Boolean
    Set referenceLayout = batch(1)
    Stats referenceLayout, Nothing, 0#, 0#, l, b, r, t, area, spread
    dx = (referenceLayout.AreaWidth - l - r) / 2#: dy = (referenceLayout.AreaHeight - b - t) / 2#
    For Each layout In batch
        For i = 1 To layout.Count
            layout.MovePlan i, dx, dy
        Next i
    Next layout
    valid = True
    For Each layout In batch
        If Not layout.Valid Then valid = False
    Next layout
    If Not valid Then
        For Each layout In batch
            For i = 1 To layout.Count
                layout.MovePlan i, -dx, -dy
            Next i
        Next layout
    End If
    AADebugWrite "Optimize", "centered=" & CStr(valid)
End Sub

Private Sub Stats(ByVal layout As AAOptimizeLayout, ByVal indices As Collection, ByVal dx As Double, _
                  ByVal dy As Double, ByRef l As Double, ByRef b As Double, ByRef r As Double, _
                  ByRef t As Double, ByRef area As Double, ByRef spread As Double)
    Dim p As AAPlacement, i As Long, j As Long, x As Double, y As Double, sx As Double, sy As Double
    Dim xx As Double, yy As Double, sumX As Double, sumY As Double
    For i = 1 To layout.Count
        Set p = layout.Placement(i)
        x = 0#: y = 0#
        If Not indices Is Nothing Then
            For j = 1 To indices.Count
                If CLng(indices(j)) = i Then x = dx: y = dy: Exit For
            Next j
        End If
        If i = 1 Then
            l = p.Footprint.LeftX + x: r = p.Footprint.RightX + x
            b = p.Footprint.BottomY + y: t = p.Footprint.TopY + y
        Else
            l = MinValue(l, p.Footprint.LeftX + x): r = MaxValue(r, p.Footprint.RightX + x)
            b = MinValue(b, p.Footprint.BottomY + y): t = MaxValue(t, p.Footprint.TopY + y)
        End If
        sx = (p.CollisionFootprint.LeftX + p.CollisionFootprint.RightX) / 2# + x
        sy = (p.CollisionFootprint.BottomY + p.CollisionFootprint.TopY) / 2# + y
        sumX = sumX + sx: sumY = sumY + sy
        xx = xx + sx * sx: yy = yy + sy * sy
    Next i
    area = (r - l) * (t - b)
    spread = xx + yy - (sumX * sumX + sumY * sumY) / layout.Count
End Sub

Private Function MovementUnits(ByVal referenceLayout As AAOptimizeLayout) As Collection
    Dim result As New Collection, indices As Collection, used() As Boolean
    Dim i As Long, j As Long, axis As Long, anchor As Double, value As Double
    Dim p As AAPlacement
    ' Baris/kolom memakai pusat kontur awal, sehingga jitter hasil pencarian
    ' individual tidak mengubah anggota kelompok gerak di pass berikutnya.
    For axis = 1 To 2
        ReDim used(1 To referenceLayout.Count)
        For i = 1 To referenceLayout.Count
            If Not used(i) Then
                Set indices = New Collection
                Set p = referenceLayout.OriginalPlacement(i)
                If axis = 1 Then
                    anchor = (p.CollisionFootprint.TopY + p.CollisionFootprint.BottomY) / 2#
                Else
                    anchor = (p.CollisionFootprint.LeftX + p.CollisionFootprint.RightX) / 2#
                End If
                For j = i To referenceLayout.Count
                    If Not used(j) Then
                        Set p = referenceLayout.OriginalPlacement(j)
                        If axis = 1 Then
                            value = (p.CollisionFootprint.TopY + p.CollisionFootprint.BottomY) / 2#
                        Else
                            value = (p.CollisionFootprint.LeftX + p.CollisionFootprint.RightX) / 2#
                        End If
                        If Abs(value - anchor) <= AA_OPTIMIZE_LINE_TOLERANCE_MM Then
                            indices.Add j: used(j) = True
                        End If
                    End If
                Next j
                If indices.Count > 1 Then result.Add indices
            End If
        Next i
    Next axis
    For i = 1 To referenceLayout.Count
        Set indices = New Collection: indices.Add i: result.Add indices
    Next i
    Set MovementUnits = result
End Function

Private Function FindRoot(ByVal page As Page, ByVal id As Long) As Shape
    Dim candidate As Shape
    For Each candidate In page.Shapes
        If CLng(candidate.StaticID) = id Then Set FindRoot = candidate: Exit Function
    Next candidate
End Function

Private Sub CheckEditable(ByVal target As Shape)
    Dim child As Shape
    If Not target.Layer.Editable Or target.Locked Then Err.Raise vbObjectError + 5368, _
        "AAOptimizeEngine", "Container atau anggotanya terkunci/tidak dapat diedit."
    If target.Type = cdrGroupShape Then
        For Each child In target.Shapes
            CheckEditable child
        Next child
    End If
End Sub

Private Function MinValue(ByVal a As Double, ByVal b As Double) As Double
    MinValue = a
    If b < a Then MinValue = b
End Function

Private Function MaxValue(ByVal a As Double, ByVal b As Double) As Double
    MaxValue = a
    If b > a Then MaxValue = b
End Function
