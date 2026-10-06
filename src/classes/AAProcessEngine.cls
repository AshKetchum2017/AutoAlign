Option Explicit

Public Sub Run(ByVal objects As AAObjectStore, ByVal modelMode As Long, _
               ByVal areaWidthText As String, ByVal areaHeightText As String, _
               ByVal horizontalGapText As String, ByVal verticalGapText As String, _
               ByVal horizontalPercent As Boolean, ByVal verticalPercent As Boolean, _
               ByRef lastProcessCompleted As Boolean, ByRef feedback As String, _
               ByRef feedbackStyle As VbMsgBoxStyle, _
               Optional ByVal layoutMode As Long = AA_MODE_MINIMUM)
    Dim settings As AASettings
    Dim doc As Document
    Dim targetPage As Page
    Dim sourceShapes() As Shape
    Dim models() As AAModel
    Dim quantities() As Long
    Dim designIndexes() As Long
    Dim layoutDesignIndexes() As Long
    Dim cutIndexes() As Long
    Dim entryCount As Long
    Dim designCount As Long
    Dim layoutDesignCount As Long
    Dim cutCount As Long
    Dim entryIndex As Long
    Dim roleIndex As Long
    Dim cutIndex As Long
    Dim sharedCutOwnerIndex As Long
    Dim planner As AAPlanner
    Dim plan As AALayoutPlan
    Dim builder As AAResultBuilder
    Dim resultSelection As ShapeRange
    Dim previousUnit As cdrUnit
    Dim previousHUnit As cdrUnit
    Dim previousVUnit As cdrUnit
    Dim originalPageWidth As Double
    Dim originalPageHeight As Double
    Dim originalCenterX As Double
    Dim originalCenterY As Double
    Dim unitsSaved As Boolean
    Dim commandGroupOpen As Boolean
    Dim pageSizeChanged As Boolean
    Dim mutationStarted As Boolean
    Dim rollbackSucceeded As Boolean
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String
    Dim operation As String
    Dim hasQuantityGroups As Boolean

    feedback = vbNullString
    feedbackStyle = vbExclamation
    On Error GoTo RunFailed
    AADebugProcessStart layoutMode, modelMode

    entryCount = objects.Count
    If entryCount = 0 Then
        feedback = "Daftarkan object Design terlebih dahulu."
        AADebugWrite "Process", "stopped: no registered objects"
        Exit Sub
    End If

    operation = "memeriksa area dan gap"
    If layoutMode <> AA_MODE_MINIMUM And layoutMode <> AA_MODE_MEDIUM Then Err.Raise _
        vbObjectError + 5329, "AAProcessEngine", "Mode penyusunan belum tersedia."
    Set settings = New AASettings
    settings.Configure modelMode, areaWidthText, areaHeightText, _
        horizontalGapText, verticalGapText, horizontalPercent, verticalPercent, _
        layoutMode = AA_MODE_MEDIUM
    AADebugSettings settings

    operation = "memeriksa dokumen dan halaman sumber"
    If Application.Documents.Count = 0 Then Err.Raise vbObjectError + 5310, _
        "AAProcessEngine", "Tidak ada dokumen aktif."
    Set doc = ActiveDocument
    Set targetPage = doc.ActivePage
    If Not (doc Is objects.SourceDocument) Then Err.Raise vbObjectError + 5311, _
        "AAProcessEngine", "Aktifkan dokumen tempat object didaftarkan."
    If targetPage.Index <> objects.SourcePageIndex Then Err.Raise vbObjectError + 5312, _
        "AAProcessEngine", "Aktifkan halaman tempat object didaftarkan."

    ReDim sourceShapes(0 To entryCount - 1)
    ReDim models(0 To entryCount - 1)
    ReDim quantities(0 To entryCount - 1)
    ReDim designIndexes(0 To entryCount - 1)
    ReDim cutIndexes(0 To entryCount - 1)
    operation = "mencari semua object terdaftar"
    For entryIndex = 0 To entryCount - 1
        objects.ValidateQuantityAt entryIndex
        If objects.QuantityGroupAt(entryIndex) > 0 Then hasQuantityGroups = True
        Set sourceShapes(entryIndex) = FindTopLevelShape(targetPage, _
            objects.StaticIDAt(entryIndex))
        If sourceShapes(entryIndex) Is Nothing Then
            If lastProcessCompleted Then Err.Raise vbObjectError + 5317, _
                "AAProcessEngine", _
                "Undo susunan sebelumnya satu kali sebelum mencoba mode lain."
            Err.Raise vbObjectError + 5313, "AAProcessEngine", _
                "Obj. ID: /" & CStr(objects.StaticIDAt(entryIndex)) & _
                " tidak ditemukan pada halaman sumber. Perbarui daftar."
        End If
        If Not sourceShapes(entryIndex).Layer.Editable Then Err.Raise _
            vbObjectError + 5314, "AAProcessEngine", _
            "Layer Obj. ID: /" & CStr(objects.StaticIDAt(entryIndex)) & _
            " tidak dapat diedit."
        quantities(entryIndex) = objects.QuantityAt(entryIndex)
        If objects.RoleAt(entryIndex) = AA_ROLE_DESIGN Then
            designIndexes(designCount) = entryIndex
            designCount = designCount + 1
        Else
            cutIndexes(cutCount) = entryIndex
            cutCount = cutCount + 1
        End If
    Next entryIndex
    If designCount = 0 Then Err.Raise vbObjectError + 5318, _
        "AAProcessEngine", "Process memerlukan sedikitnya satu Design."
    If cutCount <> 0 And cutCount <> 1 And cutCount <> designCount Then _
        Err.Raise vbObjectError + 5319, "AAProcessEngine", _
            "Jumlah Cut Line harus satu atau sama dengan jumlah Design."
    AADebugWrite "Model", "Design=" & CStr(designCount) & " Cut=" & CStr(cutCount)

    operation = "memulai CommandGroup"
    previousUnit = doc.Unit
    previousHUnit = doc.Rulers.HUnits
    previousVUnit = doc.Rulers.VUnits
    unitsSaved = True
    doc.BeginCommandGroup "Auto Align"
    commandGroupOpen = True
    doc.Unit = cdrMillimeter
    doc.Rulers.HUnits = cdrMillimeter
    doc.Rulers.VUnits = cdrMillimeter

    originalPageWidth = targetPage.SizeWidth
    originalPageHeight = targetPage.SizeHeight
    originalCenterX = targetPage.CenterX
    originalCenterY = targetPage.CenterY

    operation = "membaca kurva sumber"
    For roleIndex = 0 To designCount - 1
        entryIndex = designIndexes(roleIndex)
        cutIndex = -1
        If cutCount = 1 Then
            cutIndex = cutIndexes(0)
        ElseIf cutCount = designCount Then
            cutIndex = cutIndexes(roleIndex)
        End If
        Set models(entryIndex) = New AAModel
        If cutIndex >= 0 Then
            models(entryIndex).Prepare sourceShapes(entryIndex), _
                sourceShapes(cutIndex), settings
        Else
            models(entryIndex).Prepare sourceShapes(entryIndex), Nothing, settings
        End If
        AADebugModel entryIndex, models(entryIndex), quantities(entryIndex)
        If cutCount = 1 And roleIndex > 0 Then
            If Not models(entryIndex).CanShareCutWith( _
                   models(designIndexes(0))) Then Err.Raise _
                vbObjectError + 5326, "AAProcessEngine", _
                "Satu Cut Line bersama memerlukan bentuk Design yang sama dan selisih ukuran yang memenuhi batas 2 mm. Gunakan pasangan Cut Line masing-masing."
        End If
    Next roleIndex

    ReDim layoutDesignIndexes(0 To designCount - 1)
    For roleIndex = 0 To designCount - 1
        layoutDesignIndexes(roleIndex) = designIndexes(roleIndex)
    Next roleIndex
    layoutDesignCount = designCount
    If Not hasQuantityGroups And cutCount = 1 And designCount > 1 Then
        For roleIndex = 0 To designCount - 1
            entryIndex = designIndexes(roleIndex)
            If quantities(entryIndex) <> 0 Then
                layoutDesignIndexes(0) = entryIndex
                layoutDesignCount = 1
                Exit For
            End If
        Next roleIndex
    End If
    AADebugWrite "Model", "LayoutDesign=" & CStr(layoutDesignCount)

    operation = "merencanakan collision dan Quantity"
    Set planner = New AAPlanner
    If hasQuantityGroups Then
        operation = "merencanakan kelompok Quantity"
        Set plan = BuildQuantityGroups(objects, models, designIndexes, designCount, _
            cutIndexes, cutCount, quantities, settings, doc, layoutMode, planner)
    ElseIf layoutMode = AA_MODE_MEDIUM Then
        Set plan = planner.BuildMedium(models, layoutDesignIndexes, layoutDesignCount, _
            cutIndexes, cutCount, quantities, settings, doc)
    Else
        Set plan = planner.BuildMinimum(models, layoutDesignIndexes, layoutDesignCount, _
            cutIndexes, cutCount, quantities, settings, doc)
    End If
    AADebugPlan "Process layout", plan
    If plan.DesignCount = 0 Then
        feedback = "Tidak ada Design yang dapat ditempatkan pada area."
        AADebugWrite "Process", "no objects placed"
        GoTo CleanExit
    End If
    sharedCutOwnerIndex = -1
    If Not hasQuantityGroups And cutCount = 1 And designCount > 1 Then
        operation = "menyamakan susunan Cut Line bersama"
        Set plan = ShareCutLayout(plan, models, designIndexes, designCount, _
            quantities, planner, settings, doc, sharedCutOwnerIndex)
        AADebugPlan "Shared Cut Line expanded", plan
    End If

    ' Preset memakai ukuran sheet baku hanya selama Process. Semua koordinat
    ' hasil tetap mengacu pada pusat halaman asli yang disimpan di atas.
    If modelMode <> AA_MODEL_CUSTOM Then
        operation = "mengubah ukuran halaman sementara"
        pageSizeChanged = True
        mutationStarted = True
        targetPage.SetSize settings.PresetPageWidthMM, settings.PresetPageHeightMM
        operation = "memulihkan ukuran halaman asli"
        targetPage.SetSize originalPageWidth, originalPageHeight
        pageSizeChanged = False
    End If

    operation = "menyiapkan tool untuk pengelompokan hasil"
    If Application.ActiveTool <> cdrToolPick Then
        Application.ActiveTool = cdrToolPick
        If Application.ActiveTool <> cdrToolPick Then Err.Raise _
            vbObjectError + 5332, "AAProcessEngine", _
            "Pick Tool tidak dapat diaktifkan untuk mengelompokkan hasil."
    End If
    operation = "membuat dan mengelompokkan hasil"
    mutationStarted = True
    Set builder = New AAResultBuilder
    Set resultSelection = builder.Build(plan, sourceShapes, entryCount, settings, _
        originalCenterX, originalCenterY, sharedCutOwnerIndex)
    operation = "memilih group hasil"
    resultSelection.CreateSelection
    lastProcessCompleted = True

CleanExit:
    operation = "memulihkan unit dokumen dan ruler"
    If unitsSaved Then
        doc.Rulers.VUnits = previousVUnit
        doc.Rulers.HUnits = previousHUnit
        doc.Unit = previousUnit
        unitsSaved = False
    End If
    operation = "menutup CommandGroup"
    If commandGroupOpen Then
        doc.EndCommandGroup
        commandGroupOpen = False
    End If
    AADebugWrite "Process", "OK"
    Exit Sub

RunFailed:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    AADebugCandidateEnd
    AADebugWrite "Process", "FAILED during " & operation & "; Error " & _
        CStr(errorNumber) & " (" & errorSource & "): " & errorDescription
    On Error Resume Next
    If pageSizeChanged Then targetPage.SetSize originalPageWidth, originalPageHeight
    If unitsSaved Then
        doc.Rulers.VUnits = previousVUnit
        doc.Rulers.HUnits = previousHUnit
        doc.Unit = previousUnit
    End If
    If commandGroupOpen Then
        Err.Clear
        doc.EndCommandGroup
        If Err.Number = 0 Then
            commandGroupOpen = False
            If mutationStarted Then
                Err.Clear
                doc.Undo
                rollbackSucceeded = (Err.Number = 0)
            Else
                rollbackSucceeded = True
            End If
        End If
    End If
    On Error GoTo 0

    lastProcessCompleted = False
    feedbackStyle = vbCritical
    feedback = "Auto Align gagal saat " & operation & "." & vbCrLf & _
        "Error " & CStr(errorNumber) & " (" & errorSource & "): " & _
        errorDescription
    If commandGroupOpen Then
        feedback = feedback & vbCrLf & _
            "CommandGroup tidak berhasil ditutup; periksa Undo dan dokumen."
    ElseIf Not mutationStarted Then
        feedback = feedback & vbCrLf & "Belum ada object yang diubah."
    ElseIf rollbackSucceeded Then
        feedback = feedback & vbCrLf & "Perubahan telah dibatalkan."
    Else
        feedback = feedback & vbCrLf & _
            "Periksa hasil di dokumen sebelum melanjutkan."
    End If
End Sub

Private Function BuildQuantityGroups(ByVal objects As AAObjectStore, ByRef models() As AAModel, _
                                     ByRef designIndexes() As Long, ByVal designCount As Long, _
                                     ByRef cutIndexes() As Long, ByVal cutCount As Long, _
                                     ByRef quantities() As Long, ByVal settings As AASettings, _
                                     ByVal sourceDocument As Document, ByVal layoutMode As Long, _
                                     ByVal planner As AAPlanner) As AALayoutPlan
    Dim groups As Object, groupNumbers() As Long, keys As Variant
    Dim selectedDesigns() As Long, selectedCuts() As Long, layoutDesigns() As Long
    Dim result As AALayoutPlan, part As AALayoutPlan, balanced As AAQuantityPlanner
    Dim roleIndex As Long, i As Long, j As Long, groupNumber As Long, swap As Long
    Dim selectedCount As Long, selectedCutCount As Long, layoutCount As Long, cutOwner As Long
    Set groups = CreateObject("Scripting.Dictionary")
    For roleIndex = 0 To designCount - 1
        groupNumber = objects.QuantityGroupAt(designIndexes(roleIndex))
        If groupNumber > 0 Then groups(CStr(groupNumber)) = groupNumber
    Next roleIndex
    keys = groups.Keys
    ReDim groupNumbers(0 To groups.Count - 1)
    For i = 0 To groups.Count - 1
        groupNumbers(i) = CLng(keys(i))
    Next i
    For i = 1 To UBound(groupNumbers)
        j = i
        Do While j > 0
            If groupNumbers(j - 1) <= groupNumbers(j) Then Exit Do
            swap = groupNumbers(j - 1)
            groupNumbers(j - 1) = groupNumbers(j)
            groupNumbers(j) = swap
            j = j - 1
        Loop
    Next i
    Set result = New AALayoutPlan
    result.Configure layoutMode
    result.SetRegularPattern True
    Set balanced = New AAQuantityPlanner
    ' /n diproses menurut nomor; Quantity lama diproses sesudahnya, terpisah.
    For i = 0 To groups.Count
        groupNumber = 0
        If i < groups.Count Then groupNumber = groupNumbers(i)
        ReDim selectedDesigns(0 To designCount - 1)
        ReDim selectedCuts(0 To designCount - 1)
        selectedCount = 0: selectedCutCount = 0
        For roleIndex = 0 To designCount - 1
            If objects.QuantityGroupAt(designIndexes(roleIndex)) = groupNumber Then
                selectedDesigns(selectedCount) = designIndexes(roleIndex)
                If cutCount = designCount And cutCount > 1 Then _
                    selectedCuts(selectedCount) = cutIndexes(roleIndex)
                selectedCount = selectedCount + 1
            End If
        Next roleIndex
        If selectedCount = 0 Then GoTo NextGroup
        If cutCount = 1 Then
            selectedCuts(0) = cutIndexes(0)
            selectedCutCount = 1
        ElseIf cutCount > 1 Then
            selectedCutCount = selectedCount
        End If
        If groupNumber > 0 Then
            AADebugWrite "Quantity", "group=/" & CStr(groupNumber) & "; members=" & CStr(selectedCount)
        Else
            AADebugWrite "Quantity", "legacy quantities; members=" & CStr(selectedCount)
        End If
        cutOwner = -1
        If groupNumber > 0 Then
            Set part = balanced.BuildGroup(models, selectedDesigns, selectedCount, _
                selectedCuts, selectedCutCount, settings, sourceDocument, layoutMode, planner)
            If part.AreaCount <> 1 Then Err.Raise vbObjectError + 5334, "AAProcessEngine", _
                "Kelompok /" & CStr(groupNumber) & " harus menghasilkan tepat satu container."
        Else
            ReDim layoutDesigns(0 To selectedCount - 1)
            layoutCount = selectedCount
            For j = 0 To selectedCount - 1
                layoutDesigns(j) = selectedDesigns(j)
            Next j
            If selectedCutCount = 1 And selectedCount > 1 Then
                For j = 0 To selectedCount - 1
                    If quantities(selectedDesigns(j)) <> 0 Then
                        layoutDesigns(0) = selectedDesigns(j)
                        layoutCount = 1
                        Exit For
                    End If
                Next j
            End If
            If layoutMode = AA_MODE_MEDIUM Then
                Set part = planner.BuildMedium(models, layoutDesigns, layoutCount, selectedCuts, _
                    selectedCutCount, quantities, settings, sourceDocument)
            Else
                Set part = planner.BuildMinimum(models, layoutDesigns, layoutCount, selectedCuts, _
                    selectedCutCount, quantities, settings, sourceDocument)
            End If
            If selectedCutCount = 1 And selectedCount > 1 And part.DesignCount > 0 Then _
                Set part = ShareCutLayout(part, models, selectedDesigns, selectedCount, quantities, _
                    planner, settings, sourceDocument, cutOwner)
        End If
        AppendQuantityPlan result, part, cutOwner
NextGroup:
    Next i
    If cutCount = 1 Then SuppressDuplicateCutTemplates result
    If Not planner.FinalPlanFits(result, settings, sourceDocument) Then Err.Raise _
        vbObjectError + 5331, "AAProcessEngine", "Kelompok Quantity tidak memenuhi batas area atau mark."
    Set BuildQuantityGroups = result
End Function

Private Sub AppendQuantityPlan(ByVal result As AALayoutPlan, ByVal part As AALayoutPlan, _
                               ByVal cutOwner As Long)
    Dim placed As AAPlacement, areaBase As Long
    If part.DesignCount = 0 Then Exit Sub
    areaBase = result.AreaCount
    result.SetRegularPattern result.HasRegularPattern And part.HasRegularPattern
    For Each placed In part.Placements
        placed.SetAreaIndex areaBase + placed.AreaIndex
        ' Cut template Quantity lama hanya dicetak pada Design pemiliknya.
        If cutOwner >= 0 And placed.DesignIndex <> cutOwner Then placed.SuppressCutLine
        result.Add placed
    Next placed
End Sub

Private Sub SuppressDuplicateCutTemplates(ByVal plan As AALayoutPlan)
    Dim templates As Collection, slots As Object, reference As Object
    Dim copies As Collection
    Dim placed As AAPlacement, key As String, areaIndex As Long, i As Long
    Dim duplicate As Boolean
    Set templates = New Collection
    For areaIndex = 0 To plan.AreaCount - 1
        Set slots = CreateObject("Scripting.Dictionary")
        For Each placed In plan.Placements
            If placed.AreaIndex = areaIndex And placed.CutIndex >= 0 And placed.CreatesCutLine Then
                ' Susunan dibandingkan tanpa urutan Design; jangan menyamakan
                ' layout yang hanya punya jumlah sama atau posisi yang mendekati.
                key = CStr(placed.CutIndex) & "|" & Format$(placed.CutLeft, "0.000000000") & _
                    "|" & Format$(placed.CutTop, "0.000000000")
                If Not slots.Exists(key) Then
                    Set copies = New Collection
                    slots.Add key, copies
                End If
                Set copies = slots(key)
                copies.Add placed
            End If
        Next placed
        If slots.Count > 0 Then
            duplicate = False
            For i = 1 To templates.Count
                Set reference = templates(i)
                If SameCutSlots(slots, reference) Then
                    duplicate = True
                    Exit For
                End If
            Next i
            If duplicate Then
                For Each placed In plan.Placements
                    If placed.AreaIndex = areaIndex Then placed.SuppressCutLine
                Next placed
                AADebugWrite "Cut", "area=" & CStr(areaIndex) & "; reused template=" & CStr(i)
            Else
                templates.Add slots
                AADebugWrite "Cut", "area=" & CStr(areaIndex) & "; new template=" & CStr(templates.Count)
            End If
        End If
    Next areaIndex
End Sub

Private Function SameCutSlots(ByVal first As Object, ByVal second As Object) As Boolean
    Dim key As Variant
    Dim a As Collection, b As Collection, matched() As Boolean
    Dim i As Long, j As Long, found As Boolean
    Dim firstPlacement As AAPlacement, secondPlacement As AAPlacement
    If first.Count <> second.Count Then Exit Function
    For Each key In first.Keys
        If Not second.Exists(key) Then Exit Function
        Set a = first(key): Set b = second(key)
        If a.Count <> b.Count Then Exit Function
        ReDim matched(1 To b.Count)
        For i = 1 To a.Count
            Set firstPlacement = a(i)
            found = False
            For j = 1 To b.Count
                If Not matched(j) Then
                    Set secondPlacement = b(j)
                    ' Rotasi berbeda masih bisa identik pada bentuk simetris.
                    found = (firstPlacement.RotationDelta = secondPlacement.RotationDelta)
                    If Not found Then found = firstPlacement.CollisionFootprint.SameGeometry( _
                        secondPlacement.CollisionFootprint)
                    If found Then
                        matched(j) = True
                        found = True
                        Exit For
                    End If
                End If
            Next j
            If Not found Then Exit Function
        Next i
    Next key
    SameCutSlots = True
End Function

Private Function ShareCutLayout(ByVal sourcePlan As AALayoutPlan, _
                                ByRef models() As AAModel, _
                                ByRef designIndexes() As Long, _
                                ByVal designCount As Long, _
                                ByRef quantities() As Long, _
                                ByVal planner As AAPlanner, _
                                ByVal settings As AASettings, _
                                ByVal sourceDocument As Document, _
                                ByRef cutOwnerIndex As Long) As AALayoutPlan
    Dim template As Collection
    Dim result As AALayoutPlan
    Dim placed As AAPlacement
    Dim copiedPlacement As AAPlacement
    Dim activeCount As Long, firstIndex As Long
    Dim firstArea As Long, areaSpan As Long
    Dim roleIndex As Long, areaIndex As Long, sourceOrdinal As Long

    Set ShareCutLayout = sourcePlan
    firstIndex = -1
    For roleIndex = 0 To designCount - 1
        If quantities(designIndexes(roleIndex)) <> 0 Then
            activeCount = activeCount + 1
            If firstIndex < 0 Then firstIndex = designIndexes(roleIndex)
        End If
    Next roleIndex
    If activeCount < 2 Then Exit Function

    Set template = sourcePlan.Placements
    firstArea = sourcePlan.AreaCount
    For Each placed In template
        If placed.DesignIndex <> firstIndex Then Err.Raise vbObjectError + 5330, _
            "AAProcessEngine", "Susunan Cut Line bersama memiliki Design acuan yang berbeda."
        If placed.AreaIndex < firstArea Then firstArea = placed.AreaIndex
    Next placed
    If template.Count = 0 Then Err.Raise vbObjectError + 5330, _
        "AAProcessEngine", "Design pertama tidak memiliki susunan untuk Cut Line bersama."
    areaSpan = sourcePlan.AreaCount - firstArea

    Set result = New AALayoutPlan
    result.Configure sourcePlan.LayoutMode
    result.SetRegularPattern sourcePlan.HasRegularPattern
    For roleIndex = 0 To designCount - 1
        If quantities(designIndexes(roleIndex)) <> 0 Then
            For Each placed In template
                If sourceOrdinal = 0 Then
                    result.Add placed
                Else
                    areaIndex = sourceOrdinal * areaSpan + placed.AreaIndex - firstArea
                    Set copiedPlacement = CopySharedCutPlacement(placed, models(designIndexes(roleIndex)), _
                        designIndexes(roleIndex), areaIndex)
                    result.Add copiedPlacement
                End If
            Next placed
            sourceOrdinal = sourceOrdinal + 1
        End If
    Next roleIndex
    If Not planner.FinalPlanFits(result, settings, sourceDocument) Then Err.Raise _
        vbObjectError + 5331, "AAProcessEngine", _
        "Design tidak dapat memakai susunan Cut Line bersama pada batas area dan mark."
    cutOwnerIndex = firstIndex
    Set ShareCutLayout = result
End Function

Private Function CopySharedCutPlacement(ByVal source As AAPlacement, _
                                        ByVal model As AAModel, _
                                        ByVal designIndex As Long, _
                                        ByVal areaIndex As Long) As AAPlacement
    Dim orientation As AARotationVariant
    Dim leftX As Double, topY As Double

    If source.UsesCenter Then
        Set orientation = model.RotationVariant(source.RotationDelta)
        leftX = source.CollisionFootprint.LeftX - orientation.InsetX
        topY = source.CollisionFootprint.TopY + orientation.InsetY
        Set CopySharedCutPlacement = orientation.MakePlacement(designIndex, _
            source.CutIndex, areaIndex, leftX, topY, source.GapH, source.GapV)
    Else
        leftX = source.CutLeft - (model.DesignWidth - model.CutWidth) / 2#
        topY = source.CutTop + (model.DesignHeight - model.CutHeight) / 2#
        Set CopySharedCutPlacement = model.MakePlacement(designIndex, _
            source.CutIndex, areaIndex, leftX, topY)
    End If
End Function

Private Function FindTopLevelShape(ByVal targetPage As Page, _
                                   ByVal staticID As Long) As Shape
    Dim pageShape As Shape
    For Each pageShape In targetPage.Shapes
        If CLng(pageShape.StaticID) = staticID Then
            Set FindTopLevelShape = pageShape
            Exit Function
        End If
    Next pageShape
End Function
