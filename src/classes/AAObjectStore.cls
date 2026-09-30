Option Explicit

Private Type AAObjectEntry
    RoleLabel As String
    StaticID As Long
End Type

Private mEntries() As AAObjectEntry
Private mCount As Long
Private mSourceDocument As Document
Private mSourcePageIndex As Long

Public Property Get Count() As Long
    Count = mCount
End Property

Public Property Get SourceDocument() As Document
    Set SourceDocument = mSourceDocument
End Property

Public Property Get SourcePageIndex() As Long
    SourcePageIndex = mSourcePageIndex
End Property

Public Property Get RoleAt(ByVal index As Long) As String
    CheckIndex index
    RoleAt = mEntries(index).RoleLabel
End Property

Public Property Get StaticIDAt(ByVal index As Long) As Long
    CheckIndex index
    StaticIDAt = mEntries(index).StaticID
End Property

Public Function RegisterSelection(ByVal roleLabel As String, ByVal selectedShapes As ShapeRange, _
                                  ByVal sourceDoc As Document, ByVal sourcePage As Page) As Long
    Dim selectedShape As Shape
    Dim selectionIndex As Long

    If roleLabel <> AA_ROLE_DESIGN And roleLabel <> AA_ROLE_CUT_LINE Then Err.Raise _
        vbObjectError + 5301, "AAObjectStore", "Jenis item tidak dikenal."
    If selectedShapes Is Nothing Then Err.Raise vbObjectError + 5302, _
        "AAObjectStore", "Pilih object terlebih dahulu."
    If selectedShapes.Count = 0 Then Err.Raise vbObjectError + 5302, _
        "AAObjectStore", "Pilih object terlebih dahulu."
    If mCount > 0 Then
        If Not (sourceDoc Is mSourceDocument) Then Err.Raise vbObjectError + 5303, _
            "AAObjectStore", "Daftar hanya dapat berisi object dari satu dokumen."
        ' COM dapat memberi wrapper Page berbeda untuk halaman yang sama.
        If sourcePage.Index <> mSourcePageIndex Then Err.Raise vbObjectError + 5304, _
            "AAObjectStore", "Daftar hanya dapat berisi object dari satu halaman."
    End If

    ' Validasi seluruh seleksi sebelum menambah satu baris pun.
    For selectionIndex = 1 To selectedShapes.Count
        Set selectedShape = selectedShapes(selectionIndex)
        If Not IsTopLevelShape(sourcePage, selectedShape) Then Err.Raise _
            vbObjectError + 5305, "AAObjectStore", _
            "Pilih object langsung pada halaman, atau pilih group induknya."
        If Not selectedShape.Layer.Editable Then Err.Raise vbObjectError + 5306, _
            "AAObjectStore", "Layer object yang dipilih tidak dapat diedit."
    Next selectionIndex

    If mCount = 0 Then
        Set mSourceDocument = sourceDoc
        mSourcePageIndex = sourcePage.Index
    End If
    ' Konsisten dengan urutan seleksi yang dipakai AutoDistribute.
    For selectionIndex = selectedShapes.Count To 1 Step -1
        Set selectedShape = selectedShapes(selectionIndex)
        If Not ContainsID(CLng(selectedShape.StaticID)) Then
            AppendEntry roleLabel, CLng(selectedShape.StaticID)
            RegisterSelection = RegisterSelection + 1
        End If
    Next selectionIndex
End Function

Public Sub RemoveAt(ByVal index As Long)
    Dim i As Long

    CheckIndex index
    For i = index To mCount - 2
        mEntries(i) = mEntries(i + 1)
    Next i
    mCount = mCount - 1
    If mCount = 0 Then
        Clear
    Else
        ReDim Preserve mEntries(0 To mCount - 1)
    End If
End Sub

Public Sub Clear()
    Erase mEntries
    mCount = 0
    Set mSourceDocument = Nothing
    mSourcePageIndex = 0
End Sub

Public Sub UpdateStaticIDAt(ByVal index As Long, ByVal newStaticID As Long)
    CheckIndex index
    mEntries(index).StaticID = newStaticID
End Sub

Private Sub AppendEntry(ByVal roleLabel As String, ByVal staticID As Long)
    If mCount = 0 Then
        ReDim mEntries(0 To 0)
    Else
        ReDim Preserve mEntries(0 To mCount)
    End If
    mEntries(mCount).RoleLabel = roleLabel
    mEntries(mCount).StaticID = staticID
    mCount = mCount + 1
End Sub

Private Function ContainsID(ByVal staticID As Long) As Boolean
    Dim i As Long
    For i = 0 To mCount - 1
        If mEntries(i).StaticID = staticID Then
            ContainsID = True
            Exit Function
        End If
    Next i
End Function

Private Function IsTopLevelShape(ByVal sourcePage As Page, ByVal selectedShape As Shape) As Boolean
    Dim pageShape As Shape
    For Each pageShape In sourcePage.Shapes
        If pageShape Is selectedShape Then
            IsTopLevelShape = True
            Exit Function
        End If
    Next pageShape
End Function

Private Sub CheckIndex(ByVal index As Long)
    If index < 0 Or index >= mCount Then Err.Raise vbObjectError + 5307, _
        "AAObjectStore", "Index daftar object tidak valid."
End Sub
