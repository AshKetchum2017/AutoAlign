Option Explicit

Private Const AA_REGISTRATION_ROW_TOLERANCE_MM As Double = 1#

Private Type AAObjectEntry
    RoleLabel As String
    StaticID As Long
    Quantity As Long  ' -1 berarti kosong: isi area aktif, lalu area baru jika tidak muat.
    QuantityGroup As Long  ' 0: perilaku lama; positif: /n; -1: prefix / belum lengkap.
    RotationMode As Long
    CutPartnerID As Long  ' 0: tanpa pasangan; ID hilang tetap disimpan sampai dipilih ulang.
End Type

Private mEntries() As AAObjectEntry
Private mCount As Long
Private mSourceDocument As Document
Private mSourcePageIndex As Long

Public Property Get Count() As Long
    Count = mCount
End Property

Public Function CutPartnerIDAt(ByVal index As Long) As Long
    CheckIndex index
    CutPartnerIDAt = mEntries(index).CutPartnerID
End Function

Public Function CutIndexForID(ByVal staticID As Long) As Long
    Dim i As Long
    CutIndexForID = -1
    For i = 0 To mCount - 1
        If mEntries(i).RoleLabel = AA_ROLE_CUT_LINE And mEntries(i).StaticID = staticID Then
            CutIndexForID = i
            Exit Function
        End If
    Next i
End Function

Public Sub SetCutPartnerAt(ByVal index As Long, ByVal cutID As Long)
    CheckIndex index
    If mEntries(index).RoleLabel <> AA_ROLE_DESIGN Then Err.Raise vbObjectError + 5430, _
        "AAObjectStore", "Pasangan hanya dapat diatur pada Design."
    If cutID <> 0 Then
        If CutIndexForID(cutID) < 0 Then Err.Raise vbObjectError + 5431, _
            "AAObjectStore", "Pilih Cut Line yang terdaftar di Objects."
    End If
    mEntries(index).CutPartnerID = cutID
End Sub

Public Function CutPartnerTextAt(ByVal index As Long) As String
    Dim id As Long
    id = CutPartnerIDAt(index)
    If id = 0 Then
        CutPartnerTextAt = "Tanpa Cut Line"
    ElseIf CutIndexForID(id) < 0 Then
        CutPartnerTextAt = "Pasangan hilang /" & CStr(id)
    Else
        CutPartnerTextAt = "Cut /" & CStr(id)
    End If
End Function

' Satu kontrak untuk Process, Optimize dan Check Layout. Urutan Cut hasil
' mengikuti pasangan eksplisit Design; urutan pendaftaran tidak menentukan pasangan.
Public Sub BuildPairing(ByRef designIndexes() As Long, ByRef designCount As Long, _
                        ByRef cutIndexes() As Long, ByRef cutCount As Long)
    Dim i As Long, j As Long, id As Long, cutIndex As Long
    Dim used As Object
    designCount = 0: cutCount = 0
    Erase designIndexes: Erase cutIndexes
    If mCount = 0 Then Exit Sub
    ReDim designIndexes(0 To mCount - 1)
    ReDim cutIndexes(0 To mCount - 1)
    For i = 0 To mCount - 1
        If mEntries(i).RoleLabel = AA_ROLE_DESIGN Then
            designIndexes(designCount) = i: designCount = designCount + 1
        Else
            cutIndexes(cutCount) = i: cutCount = cutCount + 1
        End If
    Next i
    ' Optimize/Check Layout tetap mendukung Cut Line saja; Process menolaknya.
    If designCount = 0 Then Exit Sub
    Set used = CreateObject("Scripting.Dictionary")
    For j = 0 To designCount - 1
        i = designIndexes(j)
        id = mEntries(i).CutPartnerID
        If id = 0 Then
            If cutCount > 0 Then Err.Raise vbObjectError + 5432, "AAObjectStore", _
                "Design /" & CStr(mEntries(i).StaticID) & _
                " belum dipasangkan. Pilih Pasangan Cut Line, atau hapus semua Cut Line dari daftar untuk susunan tanpa pasangan."
        Else
            cutIndex = CutIndexForID(id)
            If cutIndex < 0 Then Err.Raise vbObjectError + 5433, "AAObjectStore", _
                "Pasangan Design /" & CStr(mEntries(i).StaticID) & " hilang: Cut /" & CStr(id) & _
                ". Pilih ulang pasangan atau pilih Tanpa Cut Line."
            If cutCount > 1 And used.Exists(CStr(id)) Then Err.Raise vbObjectError + 5434, _
                "AAObjectStore", "Cut /" & CStr(id) & _
                " dipakai lebih dari satu Design. Gunakan satu Cut Line bersama untuk semua Design, atau pasangan satu-per-satu."
            used(CStr(id)) = True
            If cutCount > 1 Then cutIndexes(j) = cutIndex
        End If
    Next j
    If cutCount <> 0 And cutCount <> 1 And cutCount <> designCount Then _
        Err.Raise vbObjectError + 5435, "AAObjectStore", _
            "Gunakan tanpa Cut Line, satu Cut Line bersama, atau satu Cut Line berbeda per Design. Hapus Cut Line yang tidak dipakai dari daftar."
End Sub

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

Public Property Get RotationModeAt(ByVal index As Long) As Long
    CheckIndex index
    RotationModeAt = mEntries(index).RotationMode
End Property

Public Sub SetRotationModeAt(ByVal index As Long, ByVal mode As Long)
    Dim validatedMask As Long
    CheckIndex index
    If mEntries(index).RoleLabel <> AA_ROLE_DESIGN Then Err.Raise vbObjectError + 5421, _
        "AAObjectStore", "Rotasi hanya dapat diatur pada Design."
    validatedMask = AARotationMask(mode)
    mEntries(index).RotationMode = mode
End Sub

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
    ValidateSelection selectedShapes, sourceDoc, sourcePage

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

' Corel menempatkan klik seleksi terakhir pada indeks 1. Tentukan peran
' sebelum pengurutan spasial; urutan ListBox tidak menentukan pasangan.
Public Sub RegisterSelectionPairs(ByVal selectedShapes As ShapeRange, _
                                  ByVal sourceDoc As Document, ByVal sourcePage As Page)
    Dim ids() As Long, order() As Long
    Dim designShapes As New ShapeRange
    Dim i As Long, entryIndex As Long, cutID As Long, roleLabel As String

    ValidateSelection selectedShapes, sourceDoc, sourcePage
    If selectedShapes.Count < 2 Then Err.Raise vbObjectError + 5490, _
        "AAObjectStore", "Pilih minimal dua object: Design terlebih dahulu, lalu Cut Line terakhir."
    ReDim ids(1 To selectedShapes.Count)
    ' Snapshot ID dan periksa seluruh peran sebelum mengubah daftar/pasangan.
    For i = 1 To selectedShapes.Count
        ids(i) = CLng(selectedShapes(i).StaticID)
        If i > 1 Then designShapes.Add selectedShapes(i)
        roleLabel = AA_ROLE_DESIGN
        If i = 1 Then roleLabel = AA_ROLE_CUT_LINE
        entryIndex = EntryIndexForID(ids(i))
        If entryIndex >= 0 Then
            If mEntries(entryIndex).RoleLabel <> roleLabel Then Err.Raise vbObjectError + 5491, _
                "AAObjectStore", "Object /" & CStr(ids(i)) & " sudah terdaftar sebagai " & _
                mEntries(entryIndex).RoleLabel & ". Hapus barisnya sebelum mendaftar sebagai " & roleLabel & "."
        End If
    Next i
    cutID = ids(1)
    BuildSelectionOrder designShapes, sourceDoc, order
    If mCount = 0 Then
        Set mSourceDocument = sourceDoc
        mSourcePageIndex = sourcePage.Index
    End If
    ' Design baru tetap berurutan spasial; Cut Line baru menyusul di akhir.
    For i = 1 To designShapes.Count
        If EntryIndexForID(ids(order(i) + 1)) < 0 Then AppendEntry AA_ROLE_DESIGN, ids(order(i) + 1)
    Next i
    If EntryIndexForID(cutID) < 0 Then AppendEntry AA_ROLE_CUT_LINE, cutID
    For i = 2 To selectedShapes.Count
        SetCutPartnerAt EntryIndexForID(ids(i)), cutID
    Next i
End Sub

Private Sub ValidateSelection(ByVal selectedShapes As ShapeRange, _
                              ByVal sourceDoc As Document, ByVal sourcePage As Page)
    Dim selectedShape As Shape, selectionIndex As Long

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

End Sub

Private Function EntryIndexForID(ByVal staticID As Long) As Long
    Dim i As Long
    EntryIndexForID = -1
    For i = 0 To mCount - 1
        If mEntries(i).StaticID = staticID Then
            EntryIndexForID = i
            Exit Function
        End If
    Next i
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
    mEntries(mCount).RotationMode = AA_ROTATION_AUTO
    mEntries(mCount).CutPartnerID = 0
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
