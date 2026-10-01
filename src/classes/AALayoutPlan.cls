Option Explicit

Private mPlacements As Collection
Private mLayoutMode As Long
Private mAreaCount As Long

Private Sub Class_Initialize()
    Set mPlacements = New Collection
End Sub

Public Sub Configure(ByVal layoutMode As Long)
    mLayoutMode = layoutMode
End Sub

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
