Option Explicit

Private mProductionReport As AAProductionReport

Public Property Get ProductionReport() As AAProductionReport
    Set ProductionReport = mProductionReport
End Property

Public Sub Run(ByVal objects As AAObjectStore, ByVal modelMode As Long, _
               ByVal areaWidthText As String, ByVal areaHeightText As String, _
               ByVal horizontalGapText As String, ByVal verticalGapText As String, _
               ByRef lastProcessCompleted As Boolean, ByRef feedback As String, _
               ByRef feedbackStyle As VbMsgBoxStyle, _
               Optional ByVal layoutMode As Long = AA_MODE_MINIMUM, _
               Optional ByVal customArea As AACustomArea = Nothing, _
               Optional ByVal jobPreset As AAJobPreset = Nothing)
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
    Dim cutSourceCount As Long, hasSharedCutGroups As Boolean
    Dim entryIndex As Long
    Dim roleIndex As Long
    Dim cutIndex As Long
    Dim sharedCutOwnerIndex As Long
    Dim planner As AAPlanner
    Dim plan As AALayoutPlan
    Dim builder As AAResultBuilder
    Dim resultSelection As ShapeRange
    Dim originalSelection As ShapeRange, selectionIDs() As Long, selectionCount As Long
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
    Dim cancelled As Boolean
    Dim cleanupFailed As Boolean
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String
    Dim operation As String
    Dim hasQuantityGroups As Boolean
    Dim perfStarted As Double, perfStage As String
    Dim productionDraft As AAProductionReport, startedAt As Date
    Dim cutModels As Object, cutKey As String, cutModelIndex As Long

    Set mProductionReport = Nothing
    startedAt = Now
    feedback = vbNullString
    feedbackStyle = vbExclamation
    On Error GoTo RunFailed
    AAProgressStage "Memeriksa input"
    AADebugProcessStart layoutMode, modelMode
    perfStage = "Setup"
    perfStarted = AADebugPerfBegin()

    entryCount = objects.Count
    If entryCount = 0 Then
        feedback = "Daftarkan object Design terlebih dahulu."
        AADebugWrite "Process", "stopped: no registered objects"
        AADebugStageEnd perfStage, perfStarted
        AADebugPerfFinish "stopped"
        Exit Sub
    End If

    operation = "memeriksa area dan gap"
    If layoutMode <> AA_MODE_MINIMUM And layoutMode <> AA_MODE_MEDIUM Then Err.Raise _
        vbObjectError + 5329, "AAProcessEngine", "Mode penyusunan belum tersedia."
    Set settings = New AASettings
    settings.Configure modelMode, areaWidthText, areaHeightText, _
        horizontalGapText, verticalGapText
    settings.ApplyJobDefaults jobPreset, areaWidthText, areaHeightText

    operation = "memeriksa dokumen dan halaman sumber"
    If Application.Documents.Count = 0 Then Err.Raise vbObjectError + 5310, _
        "AAProcessEngine", "Tidak ada dokumen aktif."
    Set doc = ActiveDocument
    Set targetPage = doc.ActivePage
    If Not (doc Is objects.SourceDocument) Then Err.Raise vbObjectError + 5311, _
        "AAProcessEngine", "Aktifkan dokumen tempat object didaftarkan."
    If targetPage.Index <> objects.SourcePageIndex Then Err.Raise vbObjectError + 5312, _
        "AAProcessEngine", "Aktifkan halaman tempat object didaftarkan."
    settings.ApplyCustomArea customArea, doc, targetPage.Index
    AADebugSettings settings
    Set originalSelection = ActiveSelectionRange
    selectionCount = originalSelection.Count
    If selectionCount > 0 Then
        ReDim selectionIDs(1 To selectionCount)
        For entryIndex = 1 To selectionCount
            selectionIDs(entryIndex) = CLng(originalSelection(entryIndex).StaticID)
        Next entryIndex
    End If

    ReDim sourceShapes(0 To entryCount - 1)
    ReDim models(0 To entryCount - 1)
    ReDim quantities(0 To entryCount - 1)
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
        If objects.RoleAt(entryIndex) = AA_ROLE_CUT_LINE Then cutSourceCount = cutSourceCount + 1
    Next entryIndex
    operation = "memeriksa pasangan Design dan Cut Line"
    objects.BuildProcessPairing designIndexes, designCount, cutIndexes, cutCount
    If designCount = 0 Then Err.Raise vbObjectError + 5318, _
        "AAProcessEngine", "Process memerlukan sedikitnya satu Design."
    AADebugWrite "Model", "Design=" & CStr(designCount) & " Cut=" & CStr(cutSourceCount)
    hasSharedCutGroups = HasSharedCutReferences(cutIndexes, cutCount)

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
    AADebugStageEnd perfStage, perfStarted
    perfStage = "Prepare geometry"
    perfStarted = AADebugPerfBegin()
    AAProgressStage "Membaca geometri", "Design dan Cut Line", 0, designCount
    Set cutModels = CreateObject("Scripting.Dictionary")
    For roleIndex = 0 To designCount - 1
        AAProgressStep roleIndex, designCount
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
        ' Minimum tetap 0 derajat; jangan ubah keputusan reuse template Minimum.
        If layoutMode = AA_MODE_MEDIUM Then _
            models(entryIndex).SetAllowedRotations AARotationMask(objects.RotationModeAt(entryIndex))
        AADebugWrite "Rotation", "Design ID=" & CStr(objects.StaticIDAt(entryIndex)) & _
            "; pilihan=" & AARotationText(objects.RotationModeAt(entryIndex))
        AADebugModel entryIndex, models(entryIndex), quantities(entryIndex)
        If cutIndex >= 0 Then
            cutKey = CStr(cutIndex)
            If cutModels.Exists(cutKey) Then
                cutModelIndex = CLng(cutModels(cutKey))
                If Not models(entryIndex).CanShareCutWith(models(cutModelIndex)) Then Err.Raise _
                    vbObjectError + 5326, "AAProcessEngine", _
                    "Design yang memakai Cut /" & CStr(objects.StaticIDAt(cutIndex)) & _
                    " bersama memerlukan bentuk yang sama dan selisih ukuran dalam batas 2 mm. Gunakan pasangan Cut Line masing-masing."
            Else
                cutModels.Add cutKey, entryIndex
            End If
        End If
    Next roleIndex

    If cutCount = 1 Then RestrictSharedCutRotations models, designIndexes, designCount, quantities
    If hasSharedCutGroups Then RestrictCutGroupRotations models, designIndexes, designCount, _
        cutIndexes, quantities
    ReDim layoutDesignIndexes(0 To designCount - 1)
    AADebugStageEnd perfStage, perfStarted
    AADebugMemory "After prepare geometry"
    perfStage = vbNullString
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

    AAProgressStage "Mencari susunan terbaik"
    operation = "merencanakan collision dan Quantity"
    perfStage = "Planning total"
    perfStarted = AADebugPerfBegin()
    Set planner = New AAPlanner
    If hasQuantityGroups Then
        operation = "merencanakan kelompok Quantity"
        Set plan = BuildQuantityGroups(objects, models, designIndexes, designCount, _
            cutIndexes, cutCount, quantities, settings, doc, layoutMode, planner)
    ElseIf hasSharedCutGroups Then
        Set plan = BuildLegacyCutGroups(models, designIndexes, designCount, cutIndexes, _
            quantities, settings, doc, layoutMode, planner)
    ElseIf layoutMode = AA_MODE_MEDIUM Then
        Set plan = planner.BuildMedium(models, layoutDesignIndexes, layoutDesignCount, _
            cutIndexes, cutCount, quantities, settings, doc)
    Else
        Set plan = planner.BuildMinimum(models, layoutDesignIndexes, layoutDesignCount, _
            cutIndexes, cutCount, quantities, settings, doc)
    End If
    AADebugStageEnd perfStage, perfStarted
    perfStage = vbNullString
    AADebugMemory "After planning"
    AADebugPlan "Process layout", plan
    If plan.DesignCount = 0 Then
        feedback = "Tidak ada Design yang dapat ditempatkan pada area."
        AADebugWrite "Process", "no objects placed"
        GoTo CleanExit
    End If
    sharedCutOwnerIndex = -1
    If Not hasQuantityGroups And cutCount = 1 And designCount > 1 Then
        AAProgressStage "Menyiapkan Cut Line bersama"
        operation = "menyamakan susunan Cut Line bersama"
        perfStage = "Expand shared Cut Line"
        perfStarted = AADebugPerfBegin()
        Set plan = ShareCutLayout(plan, models, designIndexes, designCount, _
            quantities, planner, settings, doc, sharedCutOwnerIndex)
        AADebugPlan "Shared Cut Line expanded", plan
        AADebugStageEnd perfStage, perfStarted
        AADebugMemory "After shared Cut Line expansion"
        perfStage = vbNullString
    End If

    ValidatePlanRotations plan, models
    AAProgressCheckCancelled
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
    Set productionDraft = New AAProductionReport
    productionDraft.Configure objects, settings, layoutMode, doc.Name, targetPage.Index, _
        startedAt, (cutCount = 1 And designCount > 1)
    mutationStarted = True
    Set builder = New AAResultBuilder
    perfStage = "Build and group results"
    perfStarted = AADebugPerfBegin()
    Set resultSelection = builder.Build(plan, sourceShapes, entryCount, settings, _
        originalCenterX, originalCenterY, sharedCutOwnerIndex, productionDraft)
    operation = "memilih group hasil"
    resultSelection.CreateSelection
    AADebugStageEnd perfStage, perfStarted
    AADebugMemory "After build and group results"
    perfStage = vbNullString
    lastProcessCompleted = True

CleanExit:
    AAProgressCheckCancelled
    AAProgressBeginCleanup
    AAProgressStage "Memulihkan dokumen"
    perfStage = "Restore document"
    perfStarted = AADebugPerfBegin()
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
    AADebugStageEnd perfStage, perfStarted
    perfStage = vbNullString
    AADebugPerfFinish "OK"
    AADebugWrite "Process", "OK"
    ' Publikasi hanya sesudah build, pemulihan unit, dan EndCommandGroup berhasil.
    If Not productionDraft Is Nothing Then
        productionDraft.Complete Now
        Set mProductionReport = productionDraft
    End If
    Exit Sub

RunFailed:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    cancelled = (errorNumber = AA_ERROR_CANCELLED)
    AAProgressBeginCleanup cancelled
    AADebugCandidateEnd
    If cancelled Then
        If Len(perfStage) > 0 Then AADebugStageEnd perfStage & " (cancelled)", perfStarted
        AADebugWrite "Process", "CANCELLED during " & operation
    Else
        If Len(perfStage) > 0 Then AADebugStageEnd perfStage & " (failed)", perfStarted
        AADebugWrite "Process", "FAILED during " & operation & "; Error " & _
            CStr(errorNumber) & " (" & errorSource & "): " & errorDescription
    End If
    On Error Resume Next
    Err.Clear
    If pageSizeChanged Then targetPage.SetSize originalPageWidth, originalPageHeight
    If Err.Number <> 0 Then cleanupFailed = True
    Err.Clear
    If unitsSaved Then
        doc.Rulers.VUnits = previousVUnit
        doc.Rulers.HUnits = previousHUnit
        doc.Unit = previousUnit
    End If
    If Err.Number <> 0 Then cleanupFailed = True
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
    If cancelled And Not doc Is Nothing Then
        Err.Clear
        RestoreCancelledSelection doc, targetPage, selectionIDs, selectionCount
        If Err.Number <> 0 Then cleanupFailed = True
    End If
    On Error GoTo 0

    If cancelled Then
        AADebugPerfFinish "CANCELLED; completed checks only"
        lastProcessCompleted = False
        feedbackStyle = vbInformation
        feedback = vbNullString
        If cleanupFailed Or commandGroupOpen Or (mutationStarted And Not rollbackSucceeded) Then
            feedbackStyle = vbCritical
            feedback = "Pemrosesan dibatalkan, tetapi pemulihan perubahan belum lengkap. Periksa Undo dan dokumen."
        End If
        Exit Sub
    End If
    AADebugPerfFinish "FAILED; completed checks only"
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

Private Sub RestoreCancelledSelection(ByVal doc As Document, ByVal targetPage As Page, _
                                      ByRef ids() As Long, ByVal count As Long)
    Dim restored As New ShapeRange, selected As Shape, i As Long
    If count = 0 Then
        doc.ClearSelection
        Exit Sub
    End If
    For i = 1 To count
        ' Undo dapat mengganti wrapper COM; cari kembali melalui StaticID.
        Set selected = FindSelectionShape(targetPage, ids(i))
        If selected Is Nothing Then Err.Raise vbObjectError + 5391, "AAProcessEngine", _
            "Seleksi awal tidak dapat dipulihkan."
        restored.Add selected
    Next i
    restored.CreateSelection
End Sub

Private Function FindSelectionShape(ByVal parent As Object, ByVal staticID As Long) As Shape
    Dim candidate As Shape, found As Shape
    For Each candidate In parent.Shapes
        If CLng(candidate.StaticID) = staticID Then
            Set FindSelectionShape = candidate
            Exit Function
        End If
        If candidate.Type = cdrGroupShape Then
            Set found = FindSelectionShape(candidate, staticID)
            If Not found Is Nothing Then
                Set FindSelectionShape = found
                Exit Function
            End If
        End If
    Next candidate
End Function

Private Function HasSharedCutReferences(ByRef cutIndexes() As Long, ByVal cutCount As Long) As Boolean
    Dim seen As Object, i As Long, key As String
    If cutCount < 2 Then Exit Function
    Set seen = CreateObject("Scripting.Dictionary")
    For i = 0 To cutCount - 1
        key = CStr(cutIndexes(i))
        If seen.Exists(key) Then
            HasSharedCutReferences = True
            Exit Function
        End If
        seen.Add key, True
    Next i
End Function

' Referensi Cut sudah sejajar dengan setiap Design; urutan kelompok mengikuti
' kemunculan pertama dalam daftar, bukan urutan key Dictionary.
Private Function GroupDesignsByCut(ByRef designIndexes() As Long, ByVal designCount As Long, _
                                   ByRef cutIndexes() As Long, ByRef cutOrder As Collection) As Object
    Dim groups As Object, members As Collection, i As Long, key As String
    Set groups = CreateObject("Scripting.Dictionary")
    Set cutOrder = New Collection
    For i = 0 To designCount - 1
        key = CStr(cutIndexes(i))
        If Not groups.Exists(key) Then
            Set members = New Collection
            groups.Add key, members
            cutOrder.Add cutIndexes(i)
        End If
        Set members = groups(key)
        members.Add designIndexes(i)
    Next i
    Set GroupDesignsByCut = groups
End Function

Private Sub RestrictCutGroupRotations(ByRef models() As AAModel, _
                                      ByRef designIndexes() As Long, ByVal designCount As Long, _
                                      ByRef cutIndexes() As Long, ByRef quantities() As Long)
    Dim groups As Object, cutOrder As Collection, members As Collection, key As Variant
    Dim selected() As Long, i As Long
    Set groups = GroupDesignsByCut(designIndexes, designCount, cutIndexes, cutOrder)
    For Each key In cutOrder
        Set members = groups(CStr(key))
        If members.Count > 1 Then
            ReDim selected(0 To members.Count - 1)
            For i = 0 To members.Count - 1
                selected(i) = CLng(members(i + 1))
            Next i
            RestrictSharedCutRotations models, selected, members.Count, quantities
        End If
    Next key
End Sub

Private Function BuildLegacyLayout(ByRef models() As AAModel, _
                                    ByRef designIndexes() As Long, ByVal designCount As Long, _
                                    ByRef cutIndexes() As Long, ByVal cutCount As Long, _
                                    ByRef quantities() As Long, ByVal settings As AASettings, _
                                    ByVal sourceDocument As Document, ByVal layoutMode As Long, _
                                    ByVal planner As AAPlanner, ByRef cutOwner As Long) As AALayoutPlan
    Dim layoutDesigns() As Long, layoutCount As Long, i As Long, result As AALayoutPlan
    ReDim layoutDesigns(0 To designCount - 1)
    layoutCount = designCount: cutOwner = -1
    For i = 0 To designCount - 1
        layoutDesigns(i) = designIndexes(i)
    Next i
    If cutCount = 1 And designCount > 1 Then
        For i = 0 To designCount - 1
            If quantities(designIndexes(i)) <> 0 Then
                layoutDesigns(0) = designIndexes(i): layoutCount = 1
                Exit For
            End If
        Next i
    End If
    If layoutMode = AA_MODE_MEDIUM Then
        Set result = planner.BuildMedium(models, layoutDesigns, layoutCount, cutIndexes, _
            cutCount, quantities, settings, sourceDocument)
    Else
        Set result = planner.BuildMinimum(models, layoutDesigns, layoutCount, cutIndexes, _
            cutCount, quantities, settings, sourceDocument)
    End If
    If cutCount = 1 And designCount > 1 And result.DesignCount > 0 Then _
        Set result = ShareCutLayout(result, models, designIndexes, designCount, quantities, _
            planner, settings, sourceDocument, cutOwner)
    Set BuildLegacyLayout = result
End Function

Private Function BuildLegacyCutGroups(ByRef models() As AAModel, _
                                       ByRef designIndexes() As Long, ByVal designCount As Long, _
                                       ByRef cutIndexes() As Long, ByRef quantities() As Long, _
                                       ByVal settings As AASettings, ByVal sourceDocument As Document, _
                                       ByVal layoutMode As Long, ByVal planner As AAPlanner) As AALayoutPlan
    Dim groups As Object, cutOrder As Collection, members As Collection, key As Variant
    Dim selected() As Long, oneCut(0 To 0) As Long, i As Long, cutOwner As Long
    Dim result As AALayoutPlan, part As AALayoutPlan
    Set groups = GroupDesignsByCut(designIndexes, designCount, cutIndexes, cutOrder)
    Set result = New AALayoutPlan
    result.Configure layoutMode: result.SetRegularPattern True
    For Each key In cutOrder
        AAProgressPulse
        Set members = groups(CStr(key))
        ReDim selected(0 To members.Count - 1)
        For i = 0 To members.Count - 1
            selected(i) = CLng(members(i + 1))
        Next i
        oneCut(0) = CLng(key)
        AADebugWrite "Pair group", "Cut index=" & CStr(key) & "; Design=" & CStr(members.Count)
        Set part = BuildLegacyLayout(models, selected, members.Count, oneCut, 1, _
            quantities, settings, sourceDocument, layoutMode, planner, cutOwner)
        AppendQuantityPlan result, part, cutOwner
    Next key
    SuppressDuplicateCutTemplates result
    If Not planner.FinalPlanFits(result, settings, sourceDocument) Then Err.Raise _
        vbObjectError + 5331, "AAProcessEngine", "Kelompok Cut Line tidak memenuhi batas area atau mark."
    Set BuildLegacyCutGroups = result
End Function

Private Function BuildQuantityGroups(ByVal objects As AAObjectStore, ByRef models() As AAModel, _
                                     ByRef designIndexes() As Long, ByVal designCount As Long, _
                                     ByRef cutIndexes() As Long, ByVal cutCount As Long, _
                                     ByRef quantities() As Long, ByVal settings As AASettings, _
                                     ByVal sourceDocument As Document, ByVal layoutMode As Long, _
                                     ByVal planner As AAPlanner) As AALayoutPlan
    Dim groups As Object, groupNumbers() As Long, keys As Variant
    Dim selectedDesigns() As Long, selectedCuts() As Long
    Dim result As AALayoutPlan, part As AALayoutPlan, balanced As AAQuantityPlanner
    Dim roleIndex As Long, i As Long, j As Long, groupNumber As Long, swap As Long
    Dim selectedCount As Long, selectedCutCount As Long, cutOwner As Long, sameCut As Boolean
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
        If groupNumber > 0 Then
            AAProgressGroup "Kelompok /" & CStr(groupNumber)
        Else
            AAProgressGroup "Quantity biasa"
        End If
        If cutCount = 1 Then
            selectedCuts(0) = cutIndexes(0)
            selectedCutCount = 1
        ElseIf cutCount > 1 Then
            selectedCutCount = selectedCount
            sameCut = True
            For j = 1 To selectedCount - 1
                If selectedCuts(j) <> selectedCuts(0) Then sameCut = False
            Next j
            If sameCut Then selectedCutCount = 1
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
            If HasSharedCutReferences(selectedCuts, selectedCutCount) Then
                Set part = BuildLegacyCutGroups(models, selectedDesigns, selectedCount, selectedCuts, _
                    quantities, settings, sourceDocument, layoutMode, planner)
            Else
                Set part = BuildLegacyLayout(models, selectedDesigns, selectedCount, selectedCuts, _
                    selectedCutCount, quantities, settings, sourceDocument, layoutMode, planner, cutOwner)
            End If
        End If
        AppendQuantityPlan result, part, cutOwner
NextGroup:
    Next i
    If cutCount > 0 Then SuppressDuplicateCutTemplates result
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
    Dim result As AAPlacement

    If source.UsesCenter Then
        Set orientation = model.RotationVariant(source.RotationDelta)
        leftX = source.CollisionFootprint.LeftX - orientation.InsetX
        topY = source.CollisionFootprint.TopY + orientation.InsetY
        Set result = orientation.MakePlacement(designIndex, _
            source.CutIndex, areaIndex, leftX, topY, source.GapH, source.GapV)
    Else
        leftX = source.CutLeft - (model.DesignWidth - model.CutWidth) / 2#
        topY = source.CutTop + (model.DesignHeight - model.CutHeight) / 2#
        Set result = model.MakePlacement(designIndex, _
            source.CutIndex, areaIndex, leftX, topY)
    End If
    result.SetMarkAreaAllowance source.UsesMarkAreaAllowance
    Set CopySharedCutPlacement = result
End Function

' Seluruh pemakai Cut Line bersama memakai irisan sudut yang sama pada run ini.
' Pilihan asli di ObjectStore tetap; Quantity 0 tidak membatasi Design aktif.
Private Sub RestrictSharedCutRotations(ByRef models() As AAModel, ByRef indexes() As Long, _
                                       ByVal count As Long, ByRef quantities() As Long)
    Dim i As Long, mask As Long
    mask = 255
    For i = 0 To count - 1
        If quantities(indexes(i)) <> 0 Then mask = mask And models(indexes(i)).RotationMask
    Next i
    If mask = 0 Then Err.Raise vbObjectError + 5423, "AAProcessEngine", _
        "Design yang memakai satu Cut Line bersama tidak memiliki sudut rotasi yang sama. " & _
        "Pilih batas rotasi yang kompatibel pada daftar Objects."
    For i = 0 To count - 1
        If quantities(indexes(i)) <> 0 Then models(indexes(i)).SetAllowedRotations mask
    Next i
    AADebugWrite "Rotation", "Cut Line bersama; effective mask=" & CStr(mask)
End Sub

Private Sub ValidatePlanRotations(ByVal plan As AALayoutPlan, ByRef models() As AAModel)
    Dim placed As AAPlacement
    For Each placed In plan.Placements
        AAProgressPulse
        If Not models(placed.DesignIndex).AllowsRotation(placed.RotationDelta) Then _
            Err.Raise vbObjectError + 5422, "AAProcessEngine", _
                "Rencana melanggar batas rotasi Design; susunan tidak diterapkan."
    Next placed
End Sub

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
