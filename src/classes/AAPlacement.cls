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
Private mForbiddenFootprint As AAGeometry
Private mRotationDelta As Double
Private mCenterX As Double, mCenterY As Double
Private mUsesCenter As Boolean
Private mContour As AAContour
Private mContourDeltaX As Double, mContourDeltaY As Double
Private mContourGeometry As AAGeometry
Private mContourVerified As Boolean
Private mGapH As Double, mGapV As Double
Private mSlotLeft As Double, mSlotTop As Double
Private mUsesContourGap As Boolean
Private mSuppressCutLine As Boolean
Private mUsesMarkAreaAllowance As Boolean

Friend Sub SetMarkAreaAllowance(ByVal allowed As Boolean)
    mUsesMarkAreaAllowance = allowed
End Sub

Friend Property Get UsesMarkAreaAllowance() As Boolean
    UsesMarkAreaAllowance = mUsesMarkAreaAllowance
End Property

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
    mContourVerified = False
End Sub

Public Sub SetForbiddenFootprint(ByVal geometry As AAGeometry)
    Set mForbiddenFootprint = geometry
End Sub

Public Property Get ForbiddenFootprint() As AAGeometry
    If mForbiddenFootprint Is Nothing Then
        Set ForbiddenFootprint = mCollisionFootprint
    Else
        Set ForbiddenFootprint = mForbiddenFootprint
    End If
End Property

Public Sub SetMediumData(ByVal angle As Double, ByVal centerX As Double, _
                         ByVal centerY As Double, ByVal contour As AAContour, _
                         ByVal gapH As Double, ByVal gapV As Double)
    mRotationDelta = angle
    mUsesContourGap = True
    mCenterX = centerX: mCenterY = centerY
    mUsesCenter = True
    Set mContour = contour
    SetContourPose centerX, centerY
    Set mContourGeometry = Nothing
    mGapH = gapH: mGapV = gapV
    mSlotLeft = mFootprint.LeftX: mSlotTop = mFootprint.TopY
End Sub

Public Sub SetShiftOrigin(ByVal leftX As Double, ByVal topY As Double)
    mSlotLeft = leftX: mSlotTop = topY
End Sub

Public Sub SetContourData(ByVal contour As AAContour, ByVal gapH As Double, _
                          ByVal gapV As Double, Optional ByVal useContourGap As Boolean = True)
    Set mContour = contour
    Set mContourGeometry = Nothing
    mGapH = gapH: mGapV = gapV
    mUsesContourGap = useContourGap
    mSlotLeft = mFootprint.LeftX: mSlotTop = mFootprint.TopY
    ' Capture inspeksi dapat memakai kurva aktual pada posisi semula.
    If Not contour Is Nothing Then SetContourPose _
        mCollisionFootprint.LeftX - contour.LeftX, mCollisionFootprint.TopY - contour.TopY
End Sub

' Pose relatif terhadap kurva saat Capture; jangan diturunkan lagi dari bbox copy.
Friend Sub SetContourPose(ByVal deltaX As Double, ByVal deltaY As Double)
    mContourDeltaX = deltaX: mContourDeltaY = deltaY
    mContourVerified = False
End Sub

Friend Sub SetContourGeometry(ByVal geometry As AAGeometry)
    Set mContourGeometry = geometry
    mContourVerified = False
End Sub

Friend Property Get ContourGeometry() As AAGeometry
    VerifyContourGeometry
    Set ContourGeometry = mContourGeometry
End Property

' Bbox saja dapat membulat, tetapi deformasi node/control harus diperiksa aktual.
' Dikerjakan sekali ketika kontur diperlukan; Translate membatalkan hasilnya.
Private Sub VerifyContourGeometry()
    Dim dx As Double, dy As Double, perfStarted As Double
    If mContourVerified Then Exit Sub
    If AADebugProfiling Then perfStarted = AADebugPerfBegin()
    If Not mContourGeometry Is Nothing Then
        dx = mContourDeltaX: dy = mContourDeltaY
        If mCollisionFootprint.TryTranslationOf(mContourGeometry, dx, dy) Then
            mContourDeltaX = dx: mContourDeltaY = dy
        Else
            Set mContour = New AAContour
            mContour.Capture mCollisionFootprint
            Set mContourGeometry = mCollisionFootprint
            mContourDeltaX = 0#: mContourDeltaY = 0#
        End If
    End If
    mContourVerified = True
    If AADebugProfiling Then AADebugPerfEnd AA_PERF_CONTOUR_POSE, perfStarted
End Sub

Friend Property Get ContourDeltaX() As Double
    VerifyContourGeometry
    ContourDeltaX = mContourDeltaX
End Property

Friend Property Get ContourDeltaY() As Double
    VerifyContourGeometry
    ContourDeltaY = mContourDeltaY
End Property

Friend Sub PreserveSourceRotation()
    mUsesCenter = False
    mRotationDelta = 0#
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
    Dim h As Double, v As Double, gap As Double
    h = mGapH: v = mGapV
    If other.GapH > h Then h = other.GapH
    If other.GapV > v Then v = other.GapV
    gap = h: If v > gap Then gap = v
    ' Bbox kurva aktual dapat berbeda tipis dari bbox kontur varian tersimpan.
    ' Max(H,V) membuat bukti ini konservatif untuk semua arah tetangga.
    If mCollisionFootprint.BoundsHaveGap(other.CollisionFootprint, gap, 0#, 0#) Then
        HasClearance = True
        Exit Function
    End If
    VerifyContourGeometry
    HasClearance = mContour.ClearOf(other.Contour, _
        mContourDeltaX, mContourDeltaY, other.ContourDeltaX, other.ContourDeltaY, h, v)
End Function

Friend Property Get Contour() As AAContour
    VerifyContourGeometry
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
    If Not mForbiddenFootprint Is Nothing Then Set mForbiddenFootprint = mForbiddenFootprint.Placed( _
        mForbiddenFootprint.LeftX + deltaX, mForbiddenFootprint.TopY + deltaY)
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
    mContourDeltaX = mContourDeltaX + deltaX
    mContourDeltaY = mContourDeltaY + deltaY
    mContourVerified = False
    mSlotLeft = mSlotLeft + deltaX
    mSlotTop = mSlotTop + deltaY
End Sub

' Snapshot metadata terpisah; geometry/contour bersifat read-only. Translate
' mengganti referensi geometry, sehingga centering tidak mengubah isi cache.
Friend Function CopyForPlan() As AAPlacement
    Dim result As AAPlacement
    Set result = New AAPlacement
    result.Configure mDesignIndex, mCutIndex, mAreaIndex, mDesignLeft, mDesignTop, _
        mCutLeft, mCutTop, mFootprint
    result.SetCollisionFootprint mCollisionFootprint
    result.SetForbiddenFootprint mForbiddenFootprint
    If mUsesCenter Then result.SetMediumData mRotationDelta, mCenterX, mCenterY, _
        mContour, mGapH, mGapV
    result.SetContourData mContour, mGapH, mGapV, mUsesContourGap
    result.SetContourPose mContourDeltaX, mContourDeltaY
    result.SetContourGeometry mContourGeometry
    result.SetShiftOrigin mSlotLeft, mSlotTop
    result.SetMarkAreaAllowance mUsesMarkAreaAllowance
    If mSuppressCutLine Then result.SuppressCutLine
    Set CopyForPlan = result
End Function

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
