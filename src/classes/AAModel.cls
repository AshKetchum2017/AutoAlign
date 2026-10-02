Option Explicit

Private Const AA_DEFAULT_ZIGZAG_VERTICAL_FACTOR As Double = 0.88
Private Const AA_MAX_DESIGN_CUT_DIFFERENCE_MM As Double = 2#
Private Const AA_DESIGN_UNDERSIZE_TOLERANCE_MM As Double = 0.2
Private Const AA_SIZE_TOLERANCE_MM As Double = 0.0001

Private mDesignShape As Shape
Private mCutShape As Shape
Private mDesignGeometry As AAGeometry
Private mDesignBoundary As AAGeometry
Private mCutGeometry As AAGeometry
Private mSettings As AASettings
Private mSharpRectangle As Boolean
Private mLayoutSharpRectangle As Boolean
Private mRectangle As Boolean

Public Sub Prepare(ByVal designShape As Shape, ByVal cutShape As Shape, _
                   ByVal settings As AASettings)
    Set mDesignShape = designShape
    Set mCutShape = cutShape
    Set mSettings = settings
    Set mDesignGeometry = New AAGeometry
    mDesignGeometry.Capture designShape
    Set mDesignBoundary = mDesignGeometry
    If designShape.Type = cdrGroupShape Then _
        Set mDesignBoundary = mDesignGeometry.OuterBoundary
    mSharpRectangle = mDesignBoundary.IsSharpRectangle
    mLayoutSharpRectangle = mSharpRectangle
    mRectangle = mSharpRectangle Or (designShape.Type = cdrRectangleShape)
    If Not mRectangle Then mRectangle = mDesignBoundary.IsRoundedRectangle
    Set mCutGeometry = Nothing
    If Not cutShape Is Nothing Then
        Set mCutGeometry = New AAGeometry
        mCutGeometry.Capture cutShape
        If Not mSharpRectangle Then ValidateCutSize
        ' Bentuk Cut Line menentukan pola susunan pasangan.
        mLayoutSharpRectangle = mCutGeometry.IsSharpRectangle
        mRectangle = mLayoutSharpRectangle Or (cutShape.Type = cdrRectangleShape)
        If Not mRectangle Then mRectangle = mCutGeometry.IsRoundedRectangle
    End If
End Sub

Private Sub ValidateCutSize()
    Dim dw As Double, dh As Double
    dw = DesignWidth - CutWidth
    dh = DesignHeight - CutHeight
    If dw < -AA_DESIGN_UNDERSIZE_TOLERANCE_MM - AA_SIZE_TOLERANCE_MM Or _
        dh < -AA_DESIGN_UNDERSIZE_TOLERANCE_MM - AA_SIZE_TOLERANCE_MM Or _
        dw > AA_MAX_DESIGN_CUT_DIFFERENCE_MM + _
        AA_SIZE_TOLERANCE_MM Or dh > AA_MAX_DESIGN_CUT_DIFFERENCE_MM + _
        AA_SIZE_TOLERANCE_MM Then Err.Raise vbObjectError + 5206, "AAModel", _
        "Pada setiap sumbu, Design boleh sama dengan Cut Line, lebih kecil maksimal 0,2 mm, atau lebih besar maksimal 2 mm."
End Sub

Public Sub GetSteps(ByVal layoutMode As Long, ByRef horizontalStep As Double, _
                    ByRef verticalStep As Double)
    Dim sharpPreset As Boolean
    Dim baseWidth As Double, baseHeight As Double
    If layoutMode <> AA_MODE_STRAIGHT And layoutMode <> AA_MODE_ZIGZAG Then Err.Raise _
        vbObjectError + 5203, "AAModel", "Mode penyusunan tidak dikenal."
    If mRectangle Then layoutMode = AA_MODE_STRAIGHT
    baseWidth = DesignWidth
    baseHeight = DesignHeight
    If Not mCutGeometry Is Nothing Then
        baseWidth = CutWidth
        baseHeight = CutHeight
    End If

    sharpPreset = mLayoutSharpRectangle And mSettings.ModelMode <> AA_MODEL_CUSTOM
    horizontalStep = baseWidth + mSettings.DefaultHorizontalGapMM
    If sharpPreset Then horizontalStep = baseWidth
    If layoutMode = AA_MODE_STRAIGHT Then
        verticalStep = baseHeight + mSettings.DefaultVerticalGapMM
    Else
        verticalStep = baseHeight * AA_DEFAULT_ZIGZAG_VERTICAL_FACTOR
    End If
    If sharpPreset Then verticalStep = baseHeight
    If mSettings.HasHorizontalGap Then
        If mSettings.HorizontalPercent Then
            horizontalStep = baseWidth * mSettings.HorizontalGap / 100#
        Else
            horizontalStep = baseWidth + mSettings.HorizontalGap
        End If
    End If
    If mSettings.HasVerticalGap Then
        If mSettings.VerticalPercent Then
            verticalStep = baseHeight * mSettings.VerticalGap / 100#
        Else
            verticalStep = baseHeight + mSettings.VerticalGap
        End If
    End If
    If horizontalStep <= 0# Or verticalStep <= 0# Then Err.Raise _
        vbObjectError + 5205, "AAModel", "Langkah penyusunan harus lebih besar dari nol."
End Sub

Public Sub GetBoundaryGaps(ByVal layoutMode As Long, ByRef horizontalGap As Double, _
                           ByRef verticalGap As Double)
    Dim hs As Double, vs As Double
    GetSteps layoutMode, hs, vs
    horizontalGap = hs - SpacingWidth
    verticalGap = vs - SpacingHeight
    ' 88% bawaan ZigZag berlaku antarbaris bulat, bukan sebagai gap antarjenis.
    If Not mSettings.HasVerticalGap Then
        verticalGap = mSettings.DefaultVerticalGapMM
        If mLayoutSharpRectangle And mSettings.ModelMode <> AA_MODEL_CUSTOM Then verticalGap = 0#
    End If
End Sub

Public Function MakePlacement(ByVal designIndex As Long, ByVal cutIndex As Long, _
                              ByVal areaIndex As Long, ByVal leftX As Double, _
                              ByVal topY As Double) As AAPlacement
    Dim result As AAPlacement
    Dim footprint As AAGeometry
    Dim cutLeft As Double
    Dim cutTop As Double

    Set footprint = New AAGeometry
    footprint.AddPlaced mDesignGeometry, leftX, topY
    If Not mCutGeometry Is Nothing Then
        cutLeft = leftX + (DesignWidth - CutWidth) / 2#
        cutTop = topY - (DesignHeight - CutHeight) / 2#
        footprint.AddPlaced mCutGeometry, cutLeft, cutTop
    End If
    Set result = New AAPlacement
    result.Configure designIndex, cutIndex, areaIndex, leftX, topY, _
        cutLeft, cutTop, footprint
    If Not mCutGeometry Is Nothing Then _
        result.SetCollisionFootprint mCutGeometry.Placed(cutLeft, cutTop)
    Set MakePlacement = result
End Function

Public Function SameDesignGeometry(ByVal other As AAModel) As Boolean
    SameDesignGeometry = mDesignGeometry.SameGeometry(other.DesignGeometry)
End Function

Public Function CanShareCutWith(ByVal other As AAModel) As Boolean
    If mSharpRectangle Or other.IsSharpRectangle Then
        CanShareCutWith = mDesignBoundary.SameGeometry(other.DesignBoundary)
    Else
        CanShareCutWith = mDesignBoundary.SameForm(other.DesignBoundary)
    End If
End Function

Friend Property Get DesignBoundary() As AAGeometry
    Set DesignBoundary = mDesignBoundary
End Property

Public Property Get IsRectangle() As Boolean
    IsRectangle = mRectangle
End Property

Friend Property Get DesignGeometry() As AAGeometry
    Set DesignGeometry = mDesignGeometry
End Property

Public Property Get IsSharpRectangle() As Boolean
    IsSharpRectangle = mSharpRectangle
End Property

Public Property Get DesignWidth() As Double
    DesignWidth = mDesignGeometry.Width
End Property

Public Property Get DesignHeight() As Double
    DesignHeight = mDesignGeometry.Height
End Property

Public Property Get FootprintWidth() As Double
    FootprintWidth = DesignWidth
    If CutWidth > FootprintWidth Then FootprintWidth = CutWidth
End Property

Public Property Get FootprintHeight() As Double
    FootprintHeight = DesignHeight
    If CutHeight > FootprintHeight Then FootprintHeight = CutHeight
End Property

Public Property Get DesignInsetX() As Double
    DesignInsetX = (FootprintWidth - DesignWidth) / 2#
End Property

Public Property Get DesignInsetY() As Double
    DesignInsetY = (FootprintHeight - DesignHeight) / 2#
End Property

Public Property Get CollisionInsetY() As Double
    CollisionInsetY = (FootprintHeight - SpacingHeight) / 2#
End Property

Private Property Get SpacingWidth() As Double
    SpacingWidth = DesignWidth
    If Not mCutGeometry Is Nothing Then SpacingWidth = CutWidth
End Property

Private Property Get SpacingHeight() As Double
    SpacingHeight = DesignHeight
    If Not mCutGeometry Is Nothing Then SpacingHeight = CutHeight
End Property

Public Property Get CutWidth() As Double
    If Not mCutGeometry Is Nothing Then CutWidth = mCutGeometry.Width
End Property

Public Property Get CutHeight() As Double
    If Not mCutGeometry Is Nothing Then CutHeight = mCutGeometry.Height
End Property
