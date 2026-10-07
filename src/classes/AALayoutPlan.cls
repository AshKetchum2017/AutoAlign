Option Explicit

Private mPlacements As Collection
Private mLayoutMode As Long
Private mAreaCount As Long
Private mHasRegularPattern As Boolean

Private Sub Class_Initialize()
    Set mPlacements = New Collection
End Sub

Public Sub Configure(ByVal layoutMode As Long)
    mLayoutMode = layoutMode
End Sub

' True hanya untuk kandidat yang dibentuk dari pola berulang yang diketahui.
' Pemusatan bersama dan slot kosong tidak mengubah pola tersebut.
Public Sub SetRegularPattern(ByVal regular As Boolean)
    mHasRegularPattern = regular
End Sub

Public Property Get HasRegularPattern() As Boolean
    HasRegularPattern = mHasRegularPattern
End Property

Public Sub Add(ByVal placement As AAPlacement)
    mPlacements.Add placement
    If placement.AreaIndex + 1 > mAreaCount Then _
        mAreaCount = placement.AreaIndex + 1
End Sub

Public Property Get Placements() As Collection
    Set Placements = mPlacements
End Property

Public Property Get DesignCount() As Long
    DesignCount = mPlacements.Count
End Property

Public Function DesignCountInArea(ByVal areaIndex As Long) As Long
    Dim placement As AAPlacement
    For Each placement In mPlacements
        If placement.AreaIndex = areaIndex Then _
            DesignCountInArea = DesignCountInArea + 1
    Next placement
End Function

Public Property Get AreaCount() As Long
    AreaCount = mAreaCount
End Property

Public Property Get LayoutMode() As Long
    LayoutMode = mLayoutMode
End Property

' Quota penuh harus cocok per Design, bukan hanya total yang sama.
Friend Function MatchesQuotas(ByRef designIndexes() As Long, ByVal designCount As Long, _
                              ByRef quantities() As Long, ByVal expected As Long) As Boolean
    Dim counts() As Long, placed As AAPlacement, i As Long, index As Long, total As Long
    If expected < 0 Or mAreaCount > 1 Or mPlacements.Count <> expected Then Exit Function
    ReDim counts(LBound(quantities) To UBound(quantities))
    For Each placed In mPlacements
        index = placed.DesignIndex
        If index < LBound(counts) Or index > UBound(counts) Then Exit Function
        counts(index) = counts(index) + 1
    Next placed
    For i = 0 To designCount - 1
        index = designIndexes(i)
        If quantities(index) < 0 Then Exit Function
        If counts(index) <> quantities(index) Then Exit Function
        total = total + quantities(index)
        counts(index) = 0
    Next i
    If total <> expected Then Exit Function
    For i = LBound(counts) To UBound(counts)
        If counts(i) <> 0 Then Exit Function
    Next i
    MatchesQuotas = True
End Function
