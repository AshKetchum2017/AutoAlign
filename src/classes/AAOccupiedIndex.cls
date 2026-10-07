Option Explicit

' Satu koleksi occupied aktif. Isi boleh bertambah; objek di dalamnya tidak
' digeser/diganti selama terikat. Pergantian koleksi atau Clear membatalkan indeks.
Private Const AA_OCCUPIED_CELL_MM As Double = 20#
Private Const AA_OCCUPIED_MIN_COUNT As Long = 32
Private Const AA_OCCUPIED_MAX_CELLS As Long = 256
Private Const AA_OCCUPIED_MAX_QUERY_CELLS As Long = 4096
Private Const AA_OCCUPIED_MAX_SORT As Long = 256
Private Const AA_OCCUPIED_PADDING_MM As Double = 0.0001

Private mSource As Collection, mBins As Object, mWide As Collection
Private mCount As Long, mCapacity As Long, mStamp As Long
Private mSeen() As Long, mMaxGap As Double
Private mLastNearby As Collection
Private mLastLeft As Long, mLastRight As Long, mLastBottom As Long, mLastTop As Long

Friend Sub Clear()
    Set mSource = Nothing: Set mBins = Nothing: Set mWide = Nothing
    Set mLastNearby = Nothing
    mCount = 0: mCapacity = 0: mStamp = 0: mMaxGap = 0#
    Erase mSeen
End Sub

Private Function CellRange(ByVal leftX As Double, ByVal bottomY As Double, _
                           ByVal rightX As Double, ByVal topY As Double, _
                           ByRef leftCell As Long, ByRef bottomCell As Long, _
                           ByRef rightCell As Long, ByRef topCell As Long) As Boolean
    Dim a As Double, b As Double, c As Double, d As Double
    If leftX > rightX Or bottomY > topY Then Exit Function
    a = Int(leftX / AA_OCCUPIED_CELL_MM): b = Int(bottomY / AA_OCCUPIED_CELL_MM)
    c = Int(rightX / AA_OCCUPIED_CELL_MM): d = Int(topY / AA_OCCUPIED_CELL_MM)
    If a < -2147483646# Or b < -2147483646# Or _
       c > 2147483646# Or d > 2147483646# Then Exit Function
    leftCell = CLng(a): bottomCell = CLng(b)
    rightCell = CLng(c): topCell = CLng(d)
    CellRange = True
End Function

Private Sub IndexPlacement(ByVal index As Long)
    Dim placed As AAPlacement, box As AAGeometry
    Dim leftCell As Long, bottomCell As Long, rightCell As Long, topCell As Long
    Dim x As Long, y As Long, key As String, bucket As Collection
    Set placed = mSource(index)
    Set box = placed.CollisionFootprint
    If placed.GapH > mMaxGap Then mMaxGap = placed.GapH
    If placed.GapV > mMaxGap Then mMaxGap = placed.GapV
    If Not CellRange(box.LeftX, box.BottomY, box.RightX, box.TopY, _
                     leftCell, bottomCell, rightCell, topCell) Then
        mWide.Add index
        Exit Sub
    End If
    If (CDbl(rightCell) - leftCell + 1#) * (CDbl(topCell) - bottomCell + 1#) > _
        AA_OCCUPIED_MAX_CELLS Then
        mWide.Add index
        Exit Sub
    End If
    For x = leftCell To rightCell
        For y = bottomCell To topCell
            key = CStr(x) & "|" & CStr(y)
            If mBins.Exists(key) Then
                Set bucket = mBins(key)
            Else
                Set bucket = New Collection
                mBins.Add key, bucket
            End If
            bucket.Add index
        Next y
    Next x
End Sub

Private Sub SyncIndex()
    Dim i As Long, previousCount As Long, building As Boolean
    building = (mBins Is Nothing)
    If building Then
        Set mBins = CreateObject("Scripting.Dictionary")
        Set mWide = New Collection
    End If
    If mSource.Count > mCapacity Then
        If mCapacity = 0 Then mCapacity = 32
        Do While mCapacity < mSource.Count
            mCapacity = mCapacity * 2
        Loop
        ReDim Preserve mSeen(1 To mCapacity)
    End If
    previousCount = mCount
    For i = mCount + 1 To mSource.Count
        IndexPlacement i
    Next i
    mCount = mSource.Count
    If building Or mCount <> previousCount Then Set mLastNearby = Nothing
    If AADebugProfiling Then AADebugOccupiedBuild building, mCount - previousCount
End Sub

' Superset dari tes bounding box planner: query diperbesar dengan gap terbesar
' kandidat/seluruh occupied. Tes gap dan overlap asli tetap dijalankan caller.
Friend Function Nearby(ByVal occupied As Collection, ByVal leftX As Double, _
                       ByVal bottomY As Double, ByVal rightX As Double, _
                       ByVal topY As Double, ByVal gap As Double) As Collection
    Dim leftCell As Long, bottomCell As Long, rightCell As Long, topCell As Long
    Dim x As Long, y As Long, item As Variant, key As String, index As Long
    Dim ids() As Long, count As Long, i As Long, j As Long, value As Long
    Dim result As Collection, bucket As Collection, margin As Double
    If mSource Is Nothing Then
        Set mSource = occupied
    ElseIf Not mSource Is occupied Then
        Clear
        Set mSource = occupied
    ElseIf occupied.Count < mCount Then
        Clear
        Set mSource = occupied
    End If
    If occupied.Count <= AA_OCCUPIED_MIN_COUNT Then GoTo AllOccupied
    If mBins Is Nothing Or mCount <> occupied.Count Then SyncIndex
    margin = gap
    If mMaxGap > margin Then margin = mMaxGap
    If margin < 0# Then margin = 0#
    margin = margin + AA_OCCUPIED_PADDING_MM
    If Not CellRange(leftX - margin, bottomY - margin, rightX + margin, topY + margin, _
                     leftCell, bottomCell, rightCell, topCell) Then GoTo AllOccupied
    If (CDbl(rightCell) - leftCell + 1#) * (CDbl(topCell) - bottomCell + 1#) > _
        AA_OCCUPIED_MAX_QUERY_CELLS Then GoTo AllOccupied
    If Not mLastNearby Is Nothing Then
        If leftCell = mLastLeft And rightCell = mLastRight And _
           bottomCell = mLastBottom And topCell = mLastTop Then
            Set Nearby = mLastNearby
            If AADebugProfiling Then AADebugOccupiedQuery occupied.Count, mLastNearby.Count, False, True
            Exit Function
        End If
    End If
    If mStamp >= 2147483646 Then
        Erase mSeen
        ReDim mSeen(1 To mCapacity)
        mStamp = 0
    End If
    mStamp = mStamp + 1
    ReDim ids(1 To AA_OCCUPIED_MAX_SORT + 1)
    For Each item In mWide
        index = CLng(item)
        count = count + 1
        If count > AA_OCCUPIED_MAX_SORT Then GoTo AllOccupied
        ids(count) = index: mSeen(index) = mStamp
    Next item
    For x = leftCell To rightCell
        For y = bottomCell To topCell
            key = CStr(x) & "|" & CStr(y)
            If mBins.Exists(key) Then
                Set bucket = mBins(key)
                For Each item In bucket
                    index = CLng(item)
                    If mSeen(index) <> mStamp Then
                        count = count + 1
                        If count > AA_OCCUPIED_MAX_SORT Then GoTo AllOccupied
                        ids(count) = index: mSeen(index) = mStamp
                    End If
                Next item
            End If
        Next y
    Next x
    ' ID menaik mempertahankan urutan occupied dan blocker pertama yang sama.
    For i = 2 To count
        value = ids(i): j = i - 1
        Do While j >= 1
            If ids(j) <= value Then Exit Do
            ids(j + 1) = ids(j): j = j - 1
        Loop
        ids(j + 1) = value
    Next i
    Set result = New Collection
    For i = 1 To count
        result.Add mSource(ids(i))
    Next i
    mLastLeft = leftCell: mLastRight = rightCell
    mLastBottom = bottomCell: mLastTop = topCell
    Set mLastNearby = result
    Set Nearby = result
    If AADebugProfiling Then AADebugOccupiedQuery occupied.Count, count, False, False
    Exit Function
AllOccupied:
    Set Nearby = occupied
    If AADebugProfiling Then AADebugOccupiedQuery occupied.Count, occupied.Count, True, False
End Function
