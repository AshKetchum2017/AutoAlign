Option Explicit

Private mStore As AAObjectStore
Private mLastProcessCompleted As Boolean
Private mCustomArea As AACustomArea
Private mProductionReport As AAProductionReport
Private mJobPreset As AAJobPreset

Public Property Get JobPreset() As AAJobPreset
    Set JobPreset = mJobPreset
End Property

Public Sub SetJobPreset(ByVal preset As AAJobPreset)
    If preset Is Nothing Then
        Set mJobPreset = Nothing
    Else
        Set mJobPreset = preset.Copy
    End If
End Sub

' Preflight pada draft; area aktif belum berubah jika margin/zona ditolak.
Public Function JobAreaDraft(ByVal preset As AAJobPreset, ByVal doc As Document) As AACustomArea
    Dim draft As AACustomArea
    preset.Validate
    If preset.ModelMode <> AA_MODEL_CUSTOM Then Exit Function
    If doc Is Nothing Then Err.Raise vbObjectError + 5483, "AAPresenter", "Buka dokumen untuk menerapkan preset Custom."
    Set draft = CustomAreaDraft(doc, doc.ActivePage.Index)
    preset.ApplyMargins draft
    Set JobAreaDraft = draft
End Function

Public Property Get ProductionReport() As AAProductionReport
    Set ProductionReport = mProductionReport
End Property

Public Function CustomAreaDraft(ByVal doc As Document, ByVal pageIndex As Long) As AACustomArea
    If Not mCustomArea Is Nothing Then
        If mCustomArea.IsSource(doc, pageIndex) Then
            Set CustomAreaDraft = mCustomArea.Copy
            Exit Function
        End If
    End If
    Set CustomAreaDraft = New AACustomArea
    CustomAreaDraft.Bind doc, pageIndex
End Function

Public Sub SetCustomArea(ByVal profile As AACustomArea)
    Set mCustomArea = profile.Copy
End Sub

Private Sub Class_Initialize()
    Set mStore = New AAObjectStore
End Sub

Public Property Get LastProcessCompleted() As Boolean
    LastProcessCompleted = mLastProcessCompleted
End Property

Public Property Get ObjectCount() As Long
    ObjectCount = mStore.Count
End Property

Public Function ObjectListText(ByVal index As Long) As String
    ObjectListText = mStore.RoleAt(index) & " | Obj. ID: /" & CStr(mStore.StaticIDAt(index))
    If mStore.RoleAt(index) = AA_ROLE_DESIGN Then
        ObjectListText = ObjectListText & " -> " & mStore.CutPartnerTextAt(index)
        If Len(mStore.QuantityTextAt(index)) > 0 Then _
            ObjectListText = ObjectListText & " | Qty: " & mStore.QuantityTextAt(index)
        ObjectListText = ObjectListText & " | Rot: " & AARotationText(mStore.RotationModeAt(index))
    End If
End Function

Public Function ObjectIDAt(ByVal index As Long) As Long
    ObjectIDAt = mStore.StaticIDAt(index)
End Function

Public Function ObjectCutPartnerIDAt(ByVal index As Long) As Long
    ObjectCutPartnerIDAt = mStore.CutPartnerIDAt(index)
End Function

Public Sub SetObjectCutPartner(ByVal index As Long, ByVal cutID As Long)
    mStore.SetCutPartnerAt index, cutID
End Sub

Public Sub SelectObjectPair(ByVal index As Long)
    Dim doc As Document, page As Page, design As Shape, cut As Shape
    Dim selected As New ShapeRange, id As Long
    If Application.Documents.Count = 0 Then Err.Raise vbObjectError + 5436, _
        "AAPresenter", "Tidak ada dokumen aktif."
    Set doc = ActiveDocument: Set page = doc.ActivePage
    If Not (doc Is mStore.SourceDocument) Or page.Index <> mStore.SourcePageIndex Then _
        Err.Raise vbObjectError + 5437, "AAPresenter", "Aktifkan dokumen dan halaman sumber."
    If mStore.RoleAt(index) <> AA_ROLE_DESIGN Then Err.Raise vbObjectError + 5438, _
        "AAPresenter", "Pilih satu baris Design pada daftar Objects."
    Set design = FindPairShape(page.Shapes, mStore.StaticIDAt(index))
    If design Is Nothing Then Err.Raise vbObjectError + 5439, _
        "AAPresenter", "Design tidak ditemukan pada halaman sumber. Perbarui daftar."
    id = mStore.CutPartnerIDAt(index)
    If id <> 0 Then
        If mStore.CutIndexForID(id) < 0 Then Err.Raise vbObjectError + 5440, _
            "AAPresenter", "Pasangan Cut Line hilang dari daftar. Pilih ulang pasangan."
        Set cut = FindPairShape(page.Shapes, id)
        If cut Is Nothing Then Err.Raise vbObjectError + 5441, _
            "AAPresenter", "Cut Line tidak ditemukan pada halaman sumber. Perbarui daftar."
        selected.Add cut
    End If
    ' Selesaikan semua resolusi sebelum mengganti seleksi dokumen.
    selected.Add design
    selected.CreateSelection
End Sub

Private Function FindPairShape(ByVal shapes As Shapes, ByVal id As Long) As Shape
    Dim item As Shape, found As Shape
    For Each item In shapes
        If CLng(item.StaticID) = id Then
            Set FindPairShape = item
            Exit Function
        End If
        If item.Type = cdrGroupShape Then
            Set found = FindPairShape(item.Shapes, id)
            If Not found Is Nothing Then
                Set FindPairShape = found
                Exit Function
            End If
        End If
    Next item
End Function

Public Function ObjectRotationModeAt(ByVal index As Long) As Long
    ObjectRotationModeAt = mStore.RotationModeAt(index)
End Function

Public Sub SetObjectRotationMode(ByVal index As Long, ByVal mode As Long)
    mStore.SetRotationModeAt index, mode
End Sub

Public Function ObjectRoleAt(ByVal index As Long) As String
    ObjectRoleAt = mStore.RoleAt(index)
End Function

Public Function ObjectQuantityTextAt(ByVal index As Long) As String
    ObjectQuantityTextAt = mStore.QuantityTextAt(index)
End Function

Public Sub SetObjectQuantityText(ByVal index As Long, ByVal valueText As String)
    mStore.SetQuantityTextAt index, valueText
End Sub

Public Sub ValidateObjectQuantity(ByVal index As Long)
    mStore.ValidateQuantityAt index
End Sub

Public Sub ClearObjects()
    mStore.Clear
    mLastProcessCompleted = False
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

Public Sub RunLayout(ByVal modelMode As Long, ByVal areaWidthText As String, _
                     ByVal areaHeightText As String, ByVal horizontalGapText As String, _
                     ByVal verticalGapText As String, ByRef feedback As String, _
                     ByRef feedbackStyle As VbMsgBoxStyle, _
                     Optional ByVal layoutMode As Long = AA_MODE_MINIMUM)
    Dim engine As AAProcessEngine

    Set engine = New AAProcessEngine
    engine.Run mStore, modelMode, areaWidthText, areaHeightText, _
        horizontalGapText, verticalGapText, _
        mLastProcessCompleted, feedback, feedbackStyle, layoutMode, mCustomArea, mJobPreset
    ' Gagal/Cancel tidak mengganti catatan sukses sebelumnya.
    If Not engine.ProductionReport Is Nothing Then Set mProductionReport = engine.ProductionReport
End Sub

Public Sub OptimizeLayout(ByVal modelMode As Long, ByVal areaWidthText As String, _
                          ByVal areaHeightText As String, ByVal horizontalGapText As String, _
                          ByVal verticalGapText As String, ByRef completed As Boolean, _
                          ByRef feedback As String, ByRef feedbackStyle As VbMsgBoxStyle, _
                          Optional ByVal layoutMode As Long = AA_MODE_MINIMUM)
    Dim engine As AAOptimizeEngine
    Set engine = New AAOptimizeEngine
    engine.Run mStore, modelMode, areaWidthText, areaHeightText, horizontalGapText, _
        verticalGapText, completed, feedback, feedbackStyle, layoutMode, mCustomArea, mJobPreset
End Sub

Public Function CheckLayout(ByVal modelMode As Long, ByVal areaWidthText As String, _
                            ByVal areaHeightText As String, ByVal horizontalGapText As String, _
                            ByVal verticalGapText As String) As AAInspectionReport
    Dim engine As AAInspectionEngine
    Set engine = New AAInspectionEngine
    Set CheckLayout = engine.Run(mStore, modelMode, areaWidthText, areaHeightText, _
        horizontalGapText, verticalGapText, mCustomArea, mJobPreset)
End Function
