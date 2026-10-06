Option Explicit

Private Const AA_REGISTRATION_ROW_TOLERANCE_MM As Double = 1#

Private Type AAObjectEntry
    RoleLabel As String
    StaticID As Long
    Quantity As Long  ' -1 berarti kosong: isi area aktif, lalu area baru jika tidak muat.
    QuantityGroup As Long  ' 0: perilaku lama; positif: /n; -1: prefix / belum lengkap.
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

Public Property Get QuantityAt(ByVal index As Long) As Long
    CheckIndex index
    QuantityAt = mEntries(index).Quantity
End Property

Public Function QuantityTextAt(ByVal index As Long) As String
    CheckIndex index
    If mEntries(index).QuantityGroup < 0 Then
        QuantityTextAt = "/"
    ElseIf mEntries(index).QuantityGroup > 0 Then
        QuantityTextAt = "/" & CStr(mEntries(index).QuantityGroup)
    ElseIf mEntries(index).Quantity >= 0 Then
        QuantityTextAt = CStr(mEntries(index).Quantity)
    End If
End Function

Public Property Get QuantityGroupAt(ByVal index As Long) As Long
    CheckIndex index
    QuantityGroupAt = mEntries(index).QuantityGroup
End Property

Public Sub ValidateQuantityAt(ByVal index As Long)
    CheckIndex index
    If mEntries(index).QuantityGroup < 0 Then Err.Raise vbObjectError + 5315, _
        "AAObjectStore", "Token / harus disertai nomor kelompok positif, misalnya /1."
End Sub

Public Sub SetQuantityTextAt(ByVal index As Long, ByVal rawText As String)
    Dim valueText As String
    Dim i As Long
    Dim grouped As Boolean, number As Long

    CheckIndex index
    If mEntries(index).RoleLabel <> AA_ROLE_DESIGN Then Err.Raise _
        vbObjectError + 5308, "AAObjectStore", "Quantity hanya berlaku untuk Design."
    valueText = Trim$(rawText)
    If Len(valueText) = 0 Then
        mEntries(index).Quantity = -1
        mEntries(index).QuantityGroup = 0
        Exit Sub
    End If
    grouped = (Left$(valueText, 1) = "/")
    If grouped Then valueText = Mid$(valueText, 2)
    ' Izinkan prefix saat mengetik; Process wajib menolak / tanpa nomor.
    If grouped And Len(valueText) = 0 Then
        mEntries(index).Quantity = -1
        mEntries(index).QuantityGroup = -1
        Exit Sub
    End If
    For i = 1 To Len(valueText)
        If Mid$(valueText, i, 1) < "0" Or Mid$(valueText, i, 1) > "9" Then _
            Err.Raise vbObjectError + 5309, "AAObjectStore", _
                "Quantity harus berupa bilangan bulat nol atau lebih, atau /n dengan nomor kelompok positif."
    Next i
    On Error GoTo InvalidQuantity
    number = CLng(valueText)
    If grouped And number = 0 Then GoTo InvalidQuantity
    If grouped Then
        mEntries(index).Quantity = -1
        mEntries(index).QuantityGroup = number
    Else
        mEntries(index).Quantity = number
        mEntries(index).QuantityGroup = 0
    End If
    Exit Sub
InvalidQuantity:
    Err.Raise vbObjectError + 5309, "AAObjectStore", _
        "Quantity terlalu besar atau tidak valid."
End Sub

Public Function RegisterSelection(ByVal roleLabel As String, ByVal selectedShapes As ShapeRange, _
                                  ByVal sourceDoc As Document, ByVal sourcePage As Page) As Long
    Dim selectedShape As Shape
    Dim selectionIndex As Long
    Dim order() As Long

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

    ' Selesaikan pembacaan posisi/pengurutan sebelum mengubah data daftar.
    BuildSelectionOrder selectedShapes, sourceDoc, order
    If mCount = 0 Then
        Set mSourceDocument = sourceDoc
        mSourcePageIndex = sourcePage.Index
    End If
    For selectionIndex = 1 To selectedShapes.Count
        Set selectedShape = selectedShapes(order(selectionIndex))
        If Not ContainsID(CLng(selectedShape.StaticID)) Then
            AppendEntry roleLabel, CLng(selectedShape.StaticID)
            RegisterSelection = RegisterSelection + 1
        End If
    Next selectionIndex
End Function

Private Sub BuildSelectionOrder(ByVal selectedShapes As ShapeRange, _
                                ByVal sourceDoc As Document, ByRef order() As Long)
    Dim leftX() As Double, centerY() As Double, staticIDs() As Long
    Dim selectedShape As Shape
    Dim x As Double, y As Double, width As Double, height As Double
    Dim sourceUnit As cdrUnit
    Dim itemCount As Long, i As Long, j As Long, candidate As Long
    Dim firstInRow As Long, lastInRow As Long, rowY As Double

    itemCount = selectedShapes.Count
    ReDim order(1 To itemCount)
    order(1) = 1
    If itemCount = 1 Then Exit Sub
    ReDim leftX(1 To itemCount)
    ReDim centerY(1 To itemCount)
    ReDim staticIDs(1 To itemCount)
    sourceUnit = sourceDoc.Unit
    For i = 1 To itemCount
        Set selectedShape = selectedShapes(i)
        selectedShape.GetBoundingBox x, y, width, height
        leftX(i) = Application.ConvertUnits(x, sourceUnit, cdrMillimeter)
        centerY(i) = Application.ConvertUnits(y + height / 2#, sourceUnit, cdrMillimeter)
        staticIDs(i) = CLng(selectedShape.StaticID)
        order(i) = i
    Next i

    ' Y tertinggi menetapkan awal tiap baris. Posisi sama memakai StaticID
    ' agar urutan tetap konsisten meskipun urutan SelectionRange berbeda.
    For i = 2 To itemCount
        candidate = order(i)
        j = i - 1
        Do While j >= 1
            If centerY(order(j)) > centerY(candidate) Then Exit Do
            If centerY(order(j)) = centerY(candidate) Then
                If staticIDs(order(j)) <= staticIDs(candidate) Then Exit Do
            End If
            order(j + 1) = order(j)
            j = j - 1
        Loop
        order(j + 1) = candidate
    Next i

    firstInRow = 1
    Do While firstInRow <= itemCount
        rowY = centerY(order(firstInRow))
        lastInRow = firstInRow
        Do While lastInRow < itemCount
            ' Cadangan 1e-9 mm hanya mengatasi pembulatan konversi tepat 1 mm.
            If rowY - centerY(order(lastInRow + 1)) > _
                AA_REGISTRATION_ROW_TOLERANCE_MM + 0.000000001 Then Exit Do
            lastInRow = lastInRow + 1
        Loop
        For i = firstInRow + 1 To lastInRow
            candidate = order(i)
            j = i - 1
            Do While j >= firstInRow
                If leftX(order(j)) < leftX(candidate) Then Exit Do
                If leftX(order(j)) = leftX(candidate) Then
                    If centerY(order(j)) > centerY(candidate) Then Exit Do
                    If centerY(order(j)) = centerY(candidate) Then
                        If staticIDs(order(j)) <= staticIDs(candidate) Then Exit Do
                    End If
                End If
                order(j + 1) = order(j)
                j = j - 1
            Loop
            order(j + 1) = candidate
        Next i
        firstInRow = lastInRow + 1
    Loop
End Sub

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

Private Sub AppendEntry(ByVal roleLabel As String, ByVal staticID As Long)
    If mCount = 0 Then
        ReDim mEntries(0 To 0)
    Else
        ReDim Preserve mEntries(0 To mCount)
    End If
    mEntries(mCount).RoleLabel = roleLabel
    mEntries(mCount).StaticID = staticID
    mEntries(mCount).Quantity = -1
    mEntries(mCount).QuantityGroup = 0
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
