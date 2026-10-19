Option Explicit

Private Const AA_INSPECTION_GROUP_GAP_MM As Double = 50#
Private Const AA_INSPECTION_EDGE_MM As Double = 0.0001
Private Const AA_INSPECTION_PAIR_MM As Double = 1#
Private mSettings As AASettings, mReport As AAInspectionReport
Private mForbidden As Collection, mOccupied As Collection, mShapeIDs As Collection
Private mIndex As AAOccupiedIndex
Private mSeenIDs As Object
Private mPageX As Double, mPageY As Double

Public Function Run(ByVal objects As AAObjectStore, ByVal modelMode As Long, _
                    ByVal areaWidthText As String, ByVal areaHeightText As String, _
                    ByVal horizontalGapText As String, ByVal verticalGapText As String) As AAInspectionReport
    Dim doc As Document, page As Page, originalSelection As ShapeRange, root As Shape
    Dim designs As New Collection, cuts As New Collection, cut As Shape
    Dim planner As AAPlanner, i As Long, unitSaved As Boolean
    Dim previousUnit As cdrUnit, previousH As cdrUnit, previousV As cdrUnit
    Dim number As Long, description As String, cancelled As Boolean, cleanupError As String
    Dim started As Double, modelName As String, h As Double, v As Double
    Set mReport = New AAInspectionReport
    Set Run = mReport
    On Error GoTo Failed
    started = AADebugPerfBegin()
    AADebugWrite "Inspect", "start"
    If objects.Count = 0 Then Err.Raise vbObjectError + 5412, "AAInspectionEngine", _
        "Daftarkan container atau objek susunan ke daftar Objects terlebih dahulu."
    If Application.Documents.Count = 0 Then Err.Raise vbObjectError + 5413, "AAInspectionEngine", _
        "Tidak ada dokumen aktif."
    Set doc = ActiveDocument: Set page = doc.ActivePage
    If Not doc Is objects.SourceDocument Or page.Index <> objects.SourcePageIndex Then _
        Err.Raise vbObjectError + 5414, "AAInspectionEngine", "Aktifkan dokumen dan halaman sumber."
    Set originalSelection = ActiveSelectionRange
    Set mSettings = New AASettings
    mSettings.Configure modelMode, areaWidthText, areaHeightText, horizontalGapText, verticalGapText
    modelName = "Custom"
    If modelMode = AA_MODEL_KISS_A Then modelName = "KissA"
    If modelMode = AA_MODEL_DIE_A Then modelName = "DieA"
    h = mSettings.DefaultHorizontalGapMM: v = mSettings.DefaultVerticalGapMM
    If mSettings.HasHorizontalGap Then h = mSettings.HorizontalGap
    If mSettings.HasVerticalGap Then v = mSettings.VerticalGap
    mReport.Bind doc, page.Index, modelName & "; area " & Format$(mSettings.AreaWidthMM, "0.###") & " x " & _
        Format$(mSettings.AreaHeightMM, "0.###") & " mm; Gap H/V " & Format$(h, "0.###") & "/" & _
        Format$(v, "0.###") & " mm (pengecualian default rectangle preset tetap berlaku)." & vbCrLf & _
        "Grid Process AA, jarak antararea 50 mm. Kolom area dipilih dari pusat container terdekat."
    previousUnit = doc.Unit: previousH = doc.Rulers.HUnits: previousV = doc.Rulers.VUnits
    unitSaved = True
    doc.Unit = cdrMillimeter
    mPageX = page.CenterX: mPageY = page.CenterY
    Set planner = New AAPlanner
    Set mForbidden = planner.ForbiddenFor(mSettings, doc)
    Set mOccupied = New Collection: Set mShapeIDs = New Collection
    Set mIndex = New AAOccupiedIndex
    Set mSeenIDs = CreateObject("Scripting.Dictionary")
    For i = 0 To objects.Count - 1
        AAProgressPulse
        Set root = mReport.Resolve(objects.StaticIDAt(i))
        If root Is Nothing Then Err.Raise vbObjectError + 5415, "AAInspectionEngine", _
            "Objek terdaftar ID " & CStr(objects.StaticIDAt(i)) & " tidak ditemukan; perbarui daftar."
        If objects.RoleAt(i) = AA_ROLE_DESIGN Then designs.Add root Else cuts.Add root
    Next i
    ' Design berpasangan hanya dibaca bounds-nya. Cut Line diperiksa sebagai
    ' acuan fisik gap/collision; satu template Cut Line bersama dibaca sekali.
    AADebugWrite "Inspect", "containers Design=" & CStr(designs.Count) & "; Cut Line=" & CStr(cuts.Count)
    For i = 1 To designs.Count
        InspectRoot designs(i), (cuts.Count = 0), True
    Next i
    For i = 1 To cuts.Count
        InspectRoot cuts(i), True, False
    Next i
    If designs.Count > 0 And cuts.Count > 0 Then
        If cuts.Count <> 1 And cuts.Count <> designs.Count Then
            mReport.AddIssue "Pasangan", 0, 0, 0, _
                "Jumlah container Cut Line harus satu atau sama dengan container Design."
            mReport.MarkIncomplete "Pemetaan pasangan belum dapat ditentukan."
        Else
            For i = 1 To designs.Count
                If cuts.Count = 1 Then Set cut = cuts(1) Else Set cut = cuts(i)
                InspectPairs designs(i), cut
            Next i
        End If
    End If
    AAProgressCheckCancelled
    mReport.Finish
    GoTo Cleanup
Failed:
    number = Err.Number: description = Err.Description
    cancelled = (number = AA_ERROR_CANCELLED)
    If cancelled Then
        mReport.MarkIncomplete "Dibatalkan oleh pengguna; daftar temuan mungkin belum lengkap."
    Else
        mReport.MarkIncomplete description
        mReport.AddIssue "Pemeriksaan", 0, 0, 0, description
    End If
Cleanup:
    AAProgressBeginCleanup cancelled
    On Error Resume Next
    Err.Clear
    If unitSaved Then doc.Unit = previousUnit
    If Err.Number <> 0 Then cleanupError = " satuan dokumen"
    Err.Clear
    If unitSaved Then doc.Rulers.HUnits = previousH
    If Err.Number <> 0 Then cleanupError = cleanupError & " ruler horizontal"
    Err.Clear
    If unitSaved Then doc.Rulers.VUnits = previousV
    If Err.Number <> 0 Then cleanupError = cleanupError & " ruler vertikal"
    Err.Clear
    If Not originalSelection Is Nothing Then
        If originalSelection.Count = 0 Then doc.ClearSelection Else originalSelection.CreateSelection
    End If
    If Err.Number <> 0 Then cleanupError = cleanupError & " seleksi"
    On Error GoTo 0
    If Len(cleanupError) > 0 Then
        mReport.MarkIncomplete "State belum pulih:" & cleanupError & "."
        mReport.AddIssue "Pemulihan", 0, 0, 0, "State belum pulih:" & cleanupError & "."
    End If
    AADebugWrite "Inspect", "complete=" & CStr(mReport.Completed) & "; issues=" & CStr(mReport.TotalIssues)
    AADebugWrite "Inspect", mReport.Summary
    AADebugStageEnd "Inspection total", started
    Set mOccupied = Nothing: Set mShapeIDs = Nothing: Set mIndex = Nothing
    Set mSeenIDs = Nothing
    Set mForbidden = Nothing: Set mSettings = Nothing
End Function

' Struktur hasil AA: container -> subgroup sumber -> instance. Group instance
' tidak dipecah lagi. Objek individual juga boleh didaftarkan langsung.
Private Function Members(ByVal root As Shape) As Collection
    Dim subgroup As Shape, member As Shape, result As New Collection
    If root.Type <> cdrGroupShape Then
        result.Add root
    Else
        For Each subgroup In root.Shapes
            If subgroup.Type = cdrGroupShape Then
                For Each member In subgroup.Shapes
                    result.Add member
                Next member
            Else
                result.Add subgroup
            End If
        Next subgroup
    End If
    If result.Count = 0 Then Err.Raise vbObjectError + 5416, "AAInspectionEngine", "Container kosong."
    Set Members = result
End Function

' Area dipilih sekali per root dari kolom grid terdekat. Origin tidak mengikuti
' bbox susunan; seluruh anggota root memakai area yang sama.
Private Function RootArea(ByVal root As Shape) As Long
    Dim x As Double, y As Double, w As Double, h As Double, column As Double
    root.GetBoundingBox x, y, w, h, False
    column = Int((x + w / 2# - mPageX) / (mSettings.AreaWidthMM + AA_INSPECTION_GROUP_GAP_MM) + 0.5)
    If column < 0# Then column = 0#
    If column > 2147483646# Then Err.Raise vbObjectError + 5417, "AAInspectionEngine", "Posisi area terlalu jauh."
    RootArea = CLng(column)
End Function

Private Function AreaLeft(ByVal areaIndex As Long) As Double
    AreaLeft = mPageX - mSettings.AreaWidthMM / 2# + _
        CDbl(areaIndex) * (mSettings.AreaWidthMM + AA_INSPECTION_GROUP_GAP_MM)
End Function

Private Function AreaBottom() As Double
    AreaBottom = mPageY - mSettings.AreaHeightMM / 2#
End Function

Private Function InsideArea(ByVal box As AAGeometry, ByVal areaIndex As Long) As Boolean
    Dim x As Double, y As Double
    x = AreaLeft(areaIndex): y = AreaBottom()
    InsideArea = (box.LeftX >= x - AA_INSPECTION_EDGE_MM And box.BottomY >= y - AA_INSPECTION_EDGE_MM And _
        box.RightX <= x + mSettings.AreaWidthMM + AA_INSPECTION_EDGE_MM And _
        box.TopY <= y + mSettings.AreaHeightMM + AA_INSPECTION_EDGE_MM)
End Function

Private Sub InspectRoot(ByVal root As Shape, ByVal checkGeometry As Boolean, ByVal designRole As Boolean)
    Dim membersList As Collection, i As Long, areaIndex As Long, number As Long, description As String
    On Error GoTo Failed
    areaIndex = RootArea(root)
    Set membersList = Members(root)
    AADebugWrite "Inspect root", "ID=" & CStr(root.StaticID) & "; type=" & CStr(root.Type) & _
        "; members=" & CStr(membersList.Count) & "; geometry=" & CStr(checkGeometry) & _
        "; design=" & CStr(designRole) & "; area=" & CStr(areaIndex + 1)
    AAProgressStage "Memeriksa susunan", "Container ID " & CStr(root.StaticID) & _
        "; area " & CStr(areaIndex + 1), 0, membersList.Count
    For i = 1 To membersList.Count
        AAProgressStep i - 1, membersList.Count
        InspectShape membersList(i), areaIndex, checkGeometry, designRole
    Next i
    AAProgressStep membersList.Count, membersList.Count
    Exit Sub
Failed:
    number = Err.Number: description = Err.Description
    If number = AA_ERROR_CANCELLED Then Err.Raise number, "AAInspectionEngine", description
    mReport.MarkIncomplete "Ada container yang tidak dapat diperiksa."
    mReport.AddIssue "Container", areaIndex + 1, CLng(root.StaticID), 0, description
End Sub

Private Sub InspectShape(ByVal target As Shape, ByVal areaIndex As Long, _
                         ByVal checkGeometry As Boolean, ByVal designRole As Boolean)
    Dim geometry As AAGeometry, localGeometry As AAGeometry, obstacle As AAGeometry
    Dim contour As AAContour, placed As AAPlacement, other As AAPlacement, nearby As Collection
    Dim h As Double, v As Double, gap As Double, id As Long, otherID As Long, mark As Long
    Dim number As Long, description As String
    On Error GoTo Failed
    id = CLng(target.StaticID)
    If mSeenIDs.Exists(CStr(id)) Then
        mReport.MarkIncomplete "Daftar memuat objek yang sama melalui beberapa container/peran."
        mReport.AddIssue "Duplikat daftar", areaIndex + 1, id, 0, _
            "Objek sudah dibaca. Jangan mendaftarkan container bersamaan dengan anggotanya."
        Exit Sub
    End If
    mSeenIDs.Add CStr(id), True
    Set geometry = New AAGeometry
    If checkGeometry Then
        geometry.Capture target
        If designRole Then Set geometry = geometry.OuterBoundary
    Else
        geometry.CaptureBounds target
    End If
    mReport.CountRead
    If Not InsideArea(geometry, areaIndex) Then mReport.AddIssue "Batas area", areaIndex + 1, id, 0, _
        "Bounds objek keluar area. Acuan kiri/bawah = " & Format$(AreaLeft(areaIndex), "0.###") & _
        "/" & Format$(AreaBottom(), "0.###") & " mm; ukuran " & _
        Format$(mSettings.AreaWidthMM, "0.###") & " x " & Format$(mSettings.AreaHeightMM, "0.###") & " mm."
    If Not checkGeometry Then Exit Sub
    Set localGeometry = geometry.Placed(geometry.LeftX - AreaLeft(areaIndex), geometry.TopY - AreaBottom())
    For Each obstacle In mForbidden
        AAProgressPulse
        mark = mark + 1
        If localGeometry.Overlaps(obstacle) Then mReport.AddIssue "Zona mark", areaIndex + 1, id, 0, _
            "Acuan objek mengenai zona mark ke-" & CStr(mark) & " menurut model area yang dipilih."
    Next obstacle
    h = mSettings.DefaultHorizontalGapMM: v = mSettings.DefaultVerticalGapMM
    If mSettings.ModelMode <> AA_MODEL_CUSTOM And geometry.IsSharpRectangle Then h = 0#: v = 0#
    If mSettings.HasHorizontalGap Then h = mSettings.HorizontalGap
    If mSettings.HasVerticalGap Then v = mSettings.VerticalGap
    Set contour = New AAContour: contour.Capture geometry
    Set placed = New AAPlacement
    placed.Configure mShapeIDs.Count + 1, -1, areaIndex, geometry.LeftX, geometry.TopY, _
        geometry.LeftX, geometry.TopY, geometry
    placed.SetContourData contour, h, v
    gap = h: If v > gap Then gap = v
    Set nearby = mIndex.Nearby(mOccupied, geometry.LeftX, geometry.BottomY, geometry.RightX, geometry.TopY, gap)
    For Each other In nearby
        AAProgressPulse
        If other.AreaIndex = areaIndex Then
            mReport.CountComparison
            otherID = CLng(mShapeIDs(other.DesignIndex))
            If geometry.Overlaps(other.CollisionFootprint) Then
                mReport.AddIssue "Collision", areaIndex + 1, id, otherID, "Kontur acuan kedua objek bertumpuk."
            ElseIf Not placed.HasClearance(other) Then
                mReport.AddIssue "Gap", areaIndex + 1, id, otherID, _
                    "Pasangan ini tidak lolos pemeriksaan gap kontur AA. Gap H/V objek = " & _
                    Format$(h, "0.###") & "/" & Format$(v, "0.###") & " mm; tetangga = " & _
                    Format$(other.GapH, "0.###") & "/" & Format$(other.GapV, "0.###") & _
                    " mm. Pada posisi diagonal dipakai gap terbesar. Ini bukan nilai jarak aktual yang diukur."
            End If
        End If
    Next other
    mShapeIDs.Add id: mOccupied.Add placed
    Exit Sub
Failed:
    number = Err.Number: description = Err.Description
    If number = AA_ERROR_CANCELLED Then Err.Raise number, "AAInspectionEngine", description
    mReport.MarkIncomplete "Ada objek atau pemeriksaan geometri yang gagal."
    mReport.AddIssue "Geometri", areaIndex + 1, id, 0, description
End Sub

Private Function RelativeCenterMatches(ByVal design As Shape, ByVal designArea As Long, _
                                       ByVal cut As Shape, ByVal cutArea As Long) As Boolean
    Dim x As Double, y As Double, w As Double, h As Double, dx As Double, dy As Double
    design.GetBoundingBox x, y, w, h, False
    dx = x + w / 2# - AreaLeft(designArea): dy = y + h / 2#
    cut.GetBoundingBox x, y, w, h, False
    dx = dx - (x + w / 2# - AreaLeft(cutArea)): dy = dy - (y + h / 2#)
    RelativeCenterMatches = (dx * dx + dy * dy <= AA_INSPECTION_PAIR_MM ^ 2 + 0.000000001)
End Function

Private Sub InspectPairs(ByVal designRoot As Shape, ByVal cutRoot As Shape)
    Dim designs As Collection, cuts As Collection, used() As Boolean
    Dim designArea As Long, cutArea As Long, i As Long, j As Long, found As Long, matches As Long
    Dim number As Long, description As String, design As Shape, cut As Shape
    On Error GoTo Failed
    designArea = RootArea(designRoot): cutArea = RootArea(cutRoot)
    Set designs = Members(designRoot): Set cuts = Members(cutRoot)
    If designs.Count <> cuts.Count Then
        mReport.MarkIncomplete "Jumlah anggota pasangan tidak cocok."
        mReport.AddIssue "Pasangan", designArea + 1, CLng(designRoot.StaticID), CLng(cutRoot.StaticID), _
            "Anggota Design=" & CStr(designs.Count) & "; Cut Line=" & CStr(cuts.Count) & _
            ". Gunakan struktur container AA atau daftarkan objek individual."
        Exit Sub
    End If
    ReDim used(1 To designs.Count)
    AAProgressStage "Memeriksa pasangan", "Container ID " & CStr(designRoot.StaticID), 0, cuts.Count
    For i = 1 To cuts.Count
        AAProgressStep i - 1, cuts.Count
        Set cut = cuts(i): matches = 0: found = 0
        For j = 1 To designs.Count
            AAProgressPulse
            If RelativeCenterMatches(designs(j), designArea, cut, cutArea) Then matches = matches + 1: found = j
        Next j
        If matches = 1 Then
            Set design = designs(found)
            If used(found) Then
                mReport.MarkIncomplete "Pasangan ambigu."
                mReport.AddIssue "Pasangan", designArea + 1, CLng(design.StaticID), CLng(cut.StaticID), _
                    "Design ini cocok dengan lebih dari satu Cut Line."
            Else
                used(found) = True
                mReport.CountPair
            End If
        Else
            mReport.MarkIncomplete "Ada pasangan yang tidak dikenali secara unik."
            mReport.AddIssue "Pasangan", designArea + 1, CLng(cut.StaticID), CLng(designRoot.StaticID), _
                "Ditemukan " & CStr(matches) & " calon Design pada pusat relatif dalam toleransi 1 mm. " & _
                "Periksa pasangan/struktur group. Toleransi ini bukan kelonggaran gap."
        End If
    Next i
    Exit Sub
Failed:
    number = Err.Number: description = Err.Description
    If number = AA_ERROR_CANCELLED Then Err.Raise number, "AAInspectionEngine", description
    mReport.MarkIncomplete "Pembacaan pasangan gagal."
    mReport.AddIssue "Pasangan", designArea + 1, CLng(designRoot.StaticID), CLng(cutRoot.StaticID), description
End Sub
