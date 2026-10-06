Option Explicit

Private mDesignIndex As Long
Private mCutIndex As Long
Private mAreaIndex As Long
Private mDesignLeft As Double
Private mDesignTop As Double
Private mCutLeft As Double
Private mCutTop As Double
Private mFootprint As AAGeometry
Private mCollisionFootprint As AAGeometry
Private mRotationDelta As Double
Private mCenterX As Double, mCenterY As Double
Private mUsesCenter As Boolean
Private mContour As AAContour
Private mGapH As Double, mGapV As Double
Private mSlotLeft As Double, mSlotTop As Double
Private mUsesContourGap As Boolean
Private mSuppressCutLine As Boolean

Public Sub SuppressCutLine()
    mSuppressCutLine = True
End Sub

Public Property Get CreatesCutLine() As Boolean
    CreatesCutLine = Not mSuppressCutLine
End Property

Public Sub SetAreaIndex(ByVal areaIndex As Long)
    mAreaIndex = areaIndex
End Sub

Public Sub Configure(ByVal designIndex As Long, ByVal cutIndex As Long, _
                     ByVal areaIndex As Long, ByVal designLeft As Double, _
                     ByVal designTop As Double, ByVal cutLeft As Double, _
                     ByVal cutTop As Double, ByVal footprint As AAGeometry)
    mDesignIndex = designIndex
    mCutIndex = cutIndex
    mAreaIndex = areaIndex
    mDesignLeft = designLeft
    mDesignTop = designTop
    mCutLeft = cutLeft
    mCutTop = cutTop
    Set mFootprint = footprint
    Set mCollisionFootprint = footprint
End Sub

Public Sub SetCollisionFootprint(ByVal geometry As AAGeometry)
    Set mCollisionFootprint = geometry
End Sub

Public Sub SetMediumData(ByVal angle As Double, ByVal centerX As Double, _
                         ByVal centerY As Double, ByVal contour As AAContour, _
                         ByVal gapH As Double, ByVal gapV As Double)
    mRotationDelta = angle
    mUsesContourGap = True
    mCenterX = centerX: mCenterY = centerY
    mUsesCenter = True
    Set mContour = contour
    mGapH = gapH: mGapV = gapV
    mSlotLeft = mFootprint.LeftX: mSlotTop = mFootprint.TopY
End Sub

Public Sub SetShiftOrigin(ByVal leftX As Double, ByVal topY As Double)
    mSlotLeft = leftX: mSlotTop = topY
End Sub

Public Sub SetContourData(ByVal contour As AAContour, ByVal gapH As Double, _
                          ByVal gapV As Double, Optional ByVal useContourGap As Boolean = True)
    Set mContour = contour
    mGapH = gapH: mGapV = gapV
    mUsesContourGap = useContourGap
    mSlotLeft = mFootprint.LeftX: mSlotTop = mFootprint.TopY
End Sub

Public Property Get UsesContourGap() As Boolean
    UsesContourGap = mUsesContourGap
End Property

Public Property Get SlotLeft() As Double
    SlotLeft = mSlotLeft
End Property

Public Property Get SlotTop() As Double
    SlotTop = mSlotTop
End Property

Public Function HasClearance(ByVal other As AAPlacement) As Boolean
    Dim h As Double, v As Double
    h = mGapH: v = mGapV
    If other.GapH > h Then h = other.GapH
    If other.GapV > v Then v = other.GapV
    HasClearance = mContour.ClearOf(other.Contour, _
        mCollisionFootprint.LeftX - mContour.LeftX, _
        mCollisionFootprint.TopY - mContour.TopY, _
        other.CollisionFootprint.LeftX - other.Contour.LeftX, _
        other.CollisionFootprint.TopY - other.Contour.TopY, h, v)
End Function

Friend Property Get Contour() As AAContour
    Set Contour = mContour
End Property

Public Property Get GapH() As Double
    GapH = mGapH
End Property

Public Property Get GapV() As Double
    GapV = mGapV
End Property

Public Property Get RotationDelta() As Double
    RotationDelta = mRotationDelta
End Property

Public Property Get UsesCenter() As Boolean
    UsesCenter = mUsesCenter
End Property

Public Property Get CenterX() As Double
    CenterX = mCenterX
End Property

Public Property Get CenterY() As Double
    CenterY = mCenterY
End Property

Public Sub Translate(ByVal deltaX As Double, ByVal deltaY As Double)
    Dim sharedFootprint As Boolean
    sharedFootprint = (mCollisionFootprint Is mFootprint)
    If Not sharedFootprint Then
        Set mCollisionFootprint = mCollisionFootprint.Placed( _
            mCollisionFootprint.LeftX + deltaX, mCollisionFootprint.TopY + deltaY)
    End If
    Set mFootprint = mFootprint.Placed(mFootprint.LeftX + deltaX, _
        mFootprint.TopY + deltaY)
    If sharedFootprint Then Set mCollisionFootprint = mFootprint
    mDesignLeft = mDesignLeft + deltaX
    mDesignTop = mDesignTop + deltaY
    mCutLeft = mCutLeft + deltaX
    mCutTop = mCutTop + deltaY
    mCenterX = mCenterX + deltaX
    mCenterY = mCenterY + deltaY
    mSlotLeft = mSlotLeft + deltaX
    mSlotTop = mSlotTop + deltaY
End Sub

Public Property Get DesignIndex() As Long
    DesignIndex = mDesignIndex
End Property

Public Property Get CutIndex() As Long
    CutIndex = mCutIndex
End Property

Public Property Get AreaIndex() As Long
    AreaIndex = mAreaIndex
End Property

Public Property Get DesignLeft() As Double
    DesignLeft = mDesignLeft
End Property

Public Property Get DesignTop() As Double
    DesignTop = mDesignTop
End Property

Public Property Get CutLeft() As Double
    CutLeft = mCutLeft
End Property

Public Property Get CutTop() As Double
    CutTop = mCutTop
End Property

Public Property Get Footprint() As AAGeometry
    Set Footprint = mFootprint
End Property

Public Property Get CollisionFootprint() As AAGeometry
    Set CollisionFootprint = mCollisionFootprint
End Property
