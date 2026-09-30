Option Explicit

Private Const AA_GROUP_GAP_MM As Double = 50#

Private mStore As AAObjectStore

Private Sub Class_Initialize()
    Set mStore = New AAObjectStore
End Sub

Public Property Get ObjectCount() As Long
    ObjectCount = mStore.Count
End Property

Public Function ObjectListText(ByVal index As Long) As String
    ObjectListText = mStore.RoleAt(index) & " | Obj. ID: /" & CStr(mStore.StaticIDAt(index))
End Function

Public Sub ClearObjects()
    mStore.Clear
End Sub

Public Sub RemoveObjectAt(ByVal index As Long)
    mStore.RemoveAt index
End Sub

Public Sub RegisterSelected(ByVal roleLabel As String, ByRef feedback As String, _
                            ByRef feedbackStyle As VbMsgBoxStyle)
    Dim doc As Document
    Dim targetPage As Page
    Dim selectedShapes As ShapeRange
    Dim addedCount As Long

    feedback = vbNullString
    feedbackStyle = vbExclamation
    On Error GoTo RegisterFailed
    If Application.Documents.Count = 0 Then Err.Raise vbObjectError + 5310, _
        "AAPresenter", "Tidak ada dokumen aktif."
    Set doc = ActiveDocument
    Set targetPage = doc.ActivePage
    Set selectedShapes = ActiveSelectionRange
    addedCount = mStore.RegisterSelection(roleLabel, selectedShapes, doc, targetPage)
    If addedCount = 0 Then
        feedback = "Tidak ada object baru yang ditambahkan."
        feedbackStyle = vbInformation
    End If
    Exit Sub

RegisterFailed:
    feedback = Err.Description
End Sub

Public Sub RunLayout(ByVal layoutMode As Long, ByRef feedback As String, _
                     ByRef feedbackStyle As VbMsgBoxStyle)
    Dim doc As Document
    Dim targetPage As Page
    Dim marker As Shape
    Dim sourceShapes() As Shape
    Dim sourceLayers() As Layer
    Dim models() As AAModel
    Dim resultGroups() As Shape
    Dim plannedMembers As Collection
    Dim resultIDs() As Long
    Dim designIndexes() As Long
    Dim cutIndexes() As Long
    Dim designCount As Long
    Dim cutCount As Long
    Dim entryCount As Long
    Dim entryIndex As Long
    Dim roleIndex As Long
    Dim anchorIndex As Long
    Dim resultSelection As ShapeRange
    Dim previousUnit As cdrUnit
    Dim previousHUnit As cdrUnit
    Dim previousVUnit As cdrUnit
    Dim unitsSaved As Boolean
    Dim commandGroupOpen As Boolean
    Dim mutationStarted As Boolean
    Dim rollbackSucceeded As Boolean
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String
    Dim operation As String

    feedback = vbNullString
    feedbackStyle = vbExclamation
    On Error GoTo RunFailed

    entryCount = mStore.Count
    If entryCount = 0 Then
        feedback = "Daftarkan object Design atau Cut Line terlebih dahulu."
        Exit Sub
    End If
    operation = "Memeriksa dokumen dan halaman sumber"
    If Application.Documents.Count = 0 Then Err.Raise vbObjectError + 5310, _
        "AAPresenter", "Tidak ada dokumen aktif."
    Set doc = ActiveDocument
    Set targetPage = doc.ActivePage
    If Not (doc Is mStore.SourceDocument) Then Err.Raise vbObjectError + 5311, _
        "AAPresenter", "Aktifkan dokumen tempat object didaftarkan."
    If targetPage.Index <> mStore.SourcePageIndex Then Err.Raise vbObjectError + 5312, _
        "AAPresenter", "Aktifkan halaman tempat object didaftarkan."

    ReDim sourceShapes(0 To entryCount - 1)
    ReDim sourceLayers(0 To entryCount - 1)
    ReDim models(0 To entryCount - 1)
    ReDim resultGroups(0 To entryCount - 1)
    ReDim resultIDs(0 To entryCount - 1)
    ReDim designIndexes(0 To entryCount - 1)
    ReDim cutIndexes(0 To entryCount - 1)

    operation = "Mencari semua object terdaftar"
    For entryIndex = 0 To entryCount - 1
        Set sourceShapes(entryIndex) = AAFindTopLevelShape(targetPage, mStore.StaticIDAt(entryIndex))
        If sourceShapes(entryIndex) Is Nothing Then Err.Raise vbObjectError + 5313, _
            "AAPresenter", "Obj. ID: /" & CStr(mStore.StaticIDAt(entryIndex)) & _
            " tidak ditemukan pada halaman sumber. Perbarui daftar."
        If Not sourceShapes(entryIndex).Layer.Editable Then Err.Raise vbObjectError + 5314, _
            "AAPresenter", "Layer Obj. ID: /" & CStr(mStore.StaticIDAt(entryIndex)) & _
            " tidak dapat diedit."
        Set sourceLayers(entryIndex) = sourceShapes(entryIndex).Layer
        If mStore.RoleAt(entryIndex) = AA_ROLE_DESIGN Then
            designIndexes(designCount) = entryIndex
            designCount = designCount + 1
        Else
            cutIndexes(cutCount) = entryIndex
            cutCount = cutCount + 1
        End If
    Next entryIndex

    operation = "Membaca unit dokumen dan ruler"
    previousUnit = doc.Unit
    previousHUnit = doc.Rulers.HUnits
    previousVUnit = doc.Rulers.VUnits
    unitsSaved = True

    operation = "Memulai CommandGroup"
    doc.BeginCommandGroup "Auto Align"
    commandGroupOpen = True

    operation = "Mengatur unit milimeter"
    doc.Unit = cdrMillimeter
    doc.Rulers.HUnits = cdrMillimeter
    doc.Rulers.VUnits = cdrMillimeter

    operation = "Merencanakan susunan setiap object"
    For entryIndex = 0 To entryCount - 1
        Set models(entryIndex) = New AAModel
        models(entryIndex).Prepare layoutMode, sourceShapes(entryIndex).SizeWidth, _
            sourceShapes(entryIndex).SizeHeight
        If models(entryIndex).PositionCount < 2 Then
            feedback = "Obj. ID: /" & CStr(mStore.StaticIDAt(entryIndex)) & _
                " tidak dapat disusun lebih dari satu kali dalam marker 320 x 470 mm."
            GoTo CleanExit
        End If
    Next entryIndex

    If designCount > 0 Then
        anchorIndex = designIndexes(0)
    Else
        anchorIndex = cutIndexes(0)
    End If
    operation = "Membuat marker"
    Set marker = sourceLayers(anchorIndex).CreateRectangle( _
        targetPage.CenterX - models(anchorIndex).MarkerWidthMM / 2#, _
        targetPage.CenterY + models(anchorIndex).MarkerHeightMM / 2#, _
        targetPage.CenterX + models(anchorIndex).MarkerWidthMM / 2#, _
        targetPage.CenterY - models(anchorIndex).MarkerHeightMM / 2#)
    mutationStarted = True

    operation = "Menyusun Design"
    For roleIndex = 0 To designCount - 1
        entryIndex = designIndexes(roleIndex)
        Set resultGroups(entryIndex) = AAArrangeShape(sourceShapes(entryIndex), marker, _
            models(entryIndex).Positions, sourceLayers(entryIndex), plannedMembers)
        AAOrderGroupMembers resultGroups(entryIndex), plannedMembers
        resultGroups(entryIndex).CenterY = targetPage.CenterY
        If roleIndex = 0 Then
            resultGroups(entryIndex).CenterX = targetPage.CenterX
        Else
            resultGroups(entryIndex).LeftX = _
                resultGroups(designIndexes(roleIndex - 1)).RightX + AA_GROUP_GAP_MM
        End If
    Next roleIndex

    operation = "Menyusun Cut Line"
    For roleIndex = 0 To cutCount - 1
        entryIndex = cutIndexes(roleIndex)
        Set resultGroups(entryIndex) = AAArrangeShape(sourceShapes(entryIndex), marker, _
            models(entryIndex).Positions, sourceLayers(entryIndex), plannedMembers)
        AAOrderGroupMembers resultGroups(entryIndex), plannedMembers
        resultGroups(entryIndex).CenterY = targetPage.CenterY
        If designCount = 0 Then
            If roleIndex = 0 Then
                resultGroups(entryIndex).CenterX = targetPage.CenterX
            Else
                resultGroups(entryIndex).LeftX = _
                    resultGroups(cutIndexes(roleIndex - 1)).RightX + AA_GROUP_GAP_MM
            End If
        ElseIf roleIndex < designCount Then
            If AAMatchesDesign(resultGroups(entryIndex), _
                resultGroups(designIndexes(roleIndex)), models(entryIndex), _
                models(designIndexes(roleIndex))) Then
                resultGroups(entryIndex).CenterX = resultGroups(designIndexes(roleIndex)).CenterX
                resultGroups(entryIndex).CenterY = resultGroups(designIndexes(roleIndex)).CenterY
            Else
                AAPlaceCutToLeft resultGroups(entryIndex), resultGroups, _
                    designIndexes, cutIndexes, roleIndex, AA_GROUP_GAP_MM
            End If
        Else
            AAPlaceCutToLeft resultGroups(entryIndex), resultGroups, _
                designIndexes, cutIndexes, roleIndex, AA_GROUP_GAP_MM
        End If
    Next roleIndex

    operation = "Menghapus marker"
    marker.Delete
    Set marker = Nothing

    operation = "Menempatkan Cut Line di atas Design"
    AAOrderRoles targetPage, resultGroups, designIndexes, designCount, _
        cutIndexes, cutCount

    operation = "Memilih seluruh group hasil"
    Set resultSelection = New ShapeRange
    For entryIndex = 0 To entryCount - 1
        resultIDs(entryIndex) = CLng(resultGroups(entryIndex).StaticID)
        resultSelection.Add resultGroups(entryIndex)
    Next entryIndex
    resultSelection.CreateSelection

CleanExit:
    operation = "Memulihkan unit dokumen dan ruler"
    If unitsSaved Then
        doc.Rulers.VUnits = previousVUnit
        doc.Rulers.HUnits = previousHUnit
        doc.Unit = previousUnit
        unitsSaved = False
    End If
    operation = "Menutup CommandGroup"
    If commandGroupOpen Then
        doc.EndCommandGroup
        commandGroupOpen = False
    End If
    If mutationStarted Then
        For entryIndex = 0 To entryCount - 1
            mStore.UpdateStaticIDAt entryIndex, resultIDs(entryIndex)
        Next entryIndex
    End If
    Exit Sub

RunFailed:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    On Error Resume Next
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

    feedbackStyle = vbCritical
    feedback = "Auto Align gagal saat " & operation & "." & vbCrLf & _
        "Error " & CStr(errorNumber) & " (" & errorSource & "): " & errorDescription
    If commandGroupOpen Then
        feedback = feedback & vbCrLf & _
            "CommandGroup tidak berhasil ditutup; periksa Undo dan dokumen."
    ElseIf Not mutationStarted Then
        feedback = feedback & vbCrLf & "Belum ada object yang diubah."
    ElseIf rollbackSucceeded Then
        feedback = feedback & vbCrLf & "Perubahan telah dibatalkan."
    Else
        feedback = feedback & vbCrLf & "Periksa hasil di dokumen sebelum melanjutkan."
    End If
End Sub

Private Function AAFindTopLevelShape(ByVal targetPage As Page, ByVal staticID As Long) As Shape
    Dim pageShape As Shape
    For Each pageShape In targetPage.Shapes
        If CLng(pageShape.StaticID) = staticID Then
            Set AAFindTopLevelShape = pageShape
            Exit Function
        End If
    Next pageShape
End Function
