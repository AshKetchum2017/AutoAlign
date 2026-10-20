Option Explicit

' Snapshot berisi nilai saja; tidak menyimpan reference Document/Shape atau form.
Private mRows As Object, mRowKeys As Collection
Private mAreas As Object, mDesignTotals As Object, mCutTotals As Object
Private mSourceIDs() As Long
Private mDesignCount As Long, mCutCount As Long
Private mContext As String, mStartedAt As Date, mFinishedAt As Date
Private mCompleted As Boolean

Private Sub Class_Initialize()
    Set mRows = CreateObject("Scripting.Dictionary")
    Set mRowKeys = New Collection
    Set mAreas = CreateObject("Scripting.Dictionary")
    Set mDesignTotals = CreateObject("Scripting.Dictionary")
    Set mCutTotals = CreateObject("Scripting.Dictionary")
End Sub

Friend Sub Configure(ByVal objects As AAObjectStore, ByVal settings As AASettings, _
                     ByVal layoutMode As Long, ByVal documentName As String, _
                     ByVal pageIndex As Long, ByVal startedAt As Date, ByVal sharedCut As Boolean)
    Dim i As Long, modelText As String, modeText As String, customText As String
    mStartedAt = startedAt
    ReDim mSourceIDs(0 To objects.Count - 1)
    For i = 0 To objects.Count - 1
        mSourceIDs(i) = objects.StaticIDAt(i)
        If objects.RoleAt(i) = AA_ROLE_DESIGN Then
            mDesignTotals.Add CStr(mSourceIDs(i)), 0&
        ElseIf objects.RoleAt(i) = AA_ROLE_CUT_LINE Then
            mCutTotals.Add CStr(mSourceIDs(i)), 0&
        End If
    Next i
    Select Case settings.ModelMode
        Case AA_MODEL_KISS_A: modelText = "Kiss Cut A3+"
        Case AA_MODEL_DIE_A: modelText = "Die Cut A3+"
        Case Else: modelText = "Custom"
    End Select
    modeText = "Minimum"
    If layoutMode = AA_MODE_MEDIUM Then modeText = "Medium"
    mContext = "Dokumen saat Process: " & documentName & "; halaman " & CStr(pageIndex) & "." & vbCrLf & _
        "Model: " & modelText & "; mode: " & modeText & "." & vbCrLf & _
        "Area penuh: " & CStr(settings.FullAreaWidthMM) & " x " & CStr(settings.FullAreaHeightMM) & " mm." & vbCrLf & _
        "Area setelah margin: " & CStr(settings.AreaWidthMM) & " x " & CStr(settings.AreaHeightMM) & " mm." & vbCrLf & _
        "Gap H: " & GapText(settings.HasHorizontalGap, settings.HorizontalGap, settings.DefaultHorizontalGapMM) & _
        "; V: " & GapText(settings.HasVerticalGap, settings.VerticalGap, settings.DefaultVerticalGapMM) & "." & vbCrLf
    If settings.ModelMode <> AA_MODEL_CUSTOM Then
        mContext = mContext & "Gap default rectangle tajam preset tetap mengikuti aturan AA." & vbCrLf
    Else
        customText = settings.CustomSummary
        If Len(customText) = 0 Then customText = "Margin 0 mm; zona terlarang: 0."
        mContext = mContext & customText & vbCrLf
    End If
    If sharedCut Then
        mContext = mContext & "Pasangan: satu Cut Line bersama. Total Cut Line menghitung objek fisik yang dibuat." & vbCrLf
    Else
        mContext = mContext & "Pasangan: mengikuti daftar sumber; tanda - berarti tanpa Cut Line." & vbCrLf
    End If
    mContext = mContext & _
        "Area adalah area hasil AA, bukan perhitungan lembar bahan atau persentase pemakaian bahan." & vbCrLf & _
        "ID Design/Cut Line adalah ID sumber saat Process, bukan ID setiap duplikat." & vbCrLf & _
        "Catatan hasil Process terakhir yang berhasil dalam sesi AA ini. Perubahan manual, Undo, " & _
        "Optimize, dan perubahan pengaturan sesudahnya tidak memperbarui catatan ini."
End Sub

Private Function GapText(ByVal supplied As Boolean, ByVal value As Double, ByVal defaultValue As Double) As String
    If supplied Then
        GapText = CStr(value) & " mm (input)"
    Else
        GapText = CStr(defaultValue) & " mm (default)"
    End If
End Function

' Dipanggil builder hanya untuk Design yang sudah dibuat dan Cut Line yang benar-benar ada.
Friend Sub RecordCreated(ByVal placement As AAPlacement, ByVal hasCutLine As Boolean)
    Dim designID As Long, cutID As Long, areaNumber As Long, key As String, row As Variant
    designID = mSourceIDs(placement.DesignIndex)
    If placement.CutIndex >= 0 Then cutID = mSourceIDs(placement.CutIndex)
    areaNumber = placement.AreaIndex + 1
    key = CStr(areaNumber) & "|" & CStr(designID) & "|" & CStr(cutID)
    If Not mRows.Exists(key) Then
        mRows.Add key, Array(areaNumber, designID, cutID, 0&)
        mRowKeys.Add key
    End If
    row = mRows(key)
    row(3) = CLng(row(3)) + 1
    mRows(key) = row
    mAreas(CStr(areaNumber)) = True
    mDesignCount = mDesignCount + 1
    mDesignTotals(CStr(designID)) = CLng(mDesignTotals(CStr(designID))) + 1
    If hasCutLine Then
        mCutCount = mCutCount + 1
        mCutTotals(CStr(cutID)) = CLng(mCutTotals(CStr(cutID))) + 1
    End If
End Sub

Friend Sub Complete(ByVal finishedAt As Date)
    mFinishedAt = finishedAt
    mCompleted = True
End Sub

Public Property Get Completed() As Boolean
    Completed = mCompleted
End Property

Public Property Get RowCount() As Long
    RowCount = mRowKeys.Count
End Property

Public Function CellText(ByVal rowIndex As Long, ByVal columnIndex As Long) As String
    Dim row As Variant
    row = mRows(CStr(mRowKeys(rowIndex)))
    CellText = CStr(row(columnIndex))
    If columnIndex = 2 And CLng(row(columnIndex)) = 0 Then CellText = "-"
End Function

Public Property Get Summary() As String
    Summary = "Design: " & CStr(mDesignCount) & " | Cut Line fisik: " & CStr(mCutCount) & _
        " | Area hasil: " & CStr(mAreas.Count)
End Property

Public Property Get Details() As String
    Dim key As Variant, seconds As Long, durationText As String
    seconds = DateDiff("s", mStartedAt, mFinishedAt)
    If seconds < 0 Then seconds = 0
    durationText = CStr(seconds \ 3600) & ":" & Format$((seconds Mod 3600) \ 60, "00") & _
        ":" & Format$(seconds Mod 60, "00")
    Details = "Process selesai: " & Format$(mFinishedAt, "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
        "Durasi Process hingga pemulihan dokumen: " & durationText & "." & vbCrLf & vbCrLf & _
        mContext & vbCrLf & vbCrLf & "Total per sumber Design:" & vbCrLf
    For Each key In mDesignTotals.Keys
        Details = Details & "Design /" & CStr(key) & ": " & CStr(mDesignTotals(key)) & vbCrLf
    Next key
    Details = Details & "Total Cut Line fisik per sumber:" & vbCrLf
    If mCutTotals.Count = 0 Then Details = Details & "Tanpa Cut Line." & vbCrLf
    For Each key In mCutTotals.Keys
        Details = Details & "Cut Line /" & CStr(key) & ": " & CStr(mCutTotals(key)) & vbCrLf
    Next key
End Property

Public Function CopyText() As String
    Dim i As Long
    CopyText = "Ringkasan Hasil Produksi" & vbCrLf & Summary & vbCrLf & vbCrLf & _
        "Area" & vbTab & "Design" & vbTab & "Cut Line" & vbTab & "Jumlah Design" & vbCrLf
    For i = 1 To RowCount
        CopyText = CopyText & CellText(i, 0) & vbTab & CellText(i, 1) & vbTab & _
            CellText(i, 2) & vbTab & CellText(i, 3) & vbCrLf
    Next i
    CopyText = CopyText & vbCrLf & Details
End Function
