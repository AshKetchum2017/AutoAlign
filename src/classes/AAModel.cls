Option Explicit

Private Const AA_DEFAULT_ZIGZAG_VERTICAL_FACTOR As Double = 0.88
Private Const AA_MAX_DESIGN_CUT_DIFFERENCE_MM As Double = 2#
Private Const AA_DESIGN_UNDERSIZE_TOLERANCE_MM As Double = 0.2
Private Const AA_SIZE_TOLERANCE_MM As Double = 0.0001
Private Const AA_CONTACT_TOLERANCE_MM As Double = 0.01
Private Const AA_CONTACT_SCAN_DIVISIONS As Long = 64

Private mDesignShape As Shape
Private mCutShape As Shape
Private mDesignGeometry As AAGeometry
Private mDesignBoundary As AAGeometry
Private mCutGeometry As AAGeometry
Private mSettings As AASettings
Private mSharpRectangle As Boolean
Private mLayoutSharpRectangle As Boolean
Private mRectangle As Boolean
Private mTriangle As Boolean
Private mVariants As Object
Private mMediumDesignGeometry As AAGeometry
Private mMinimumSteps As Object

Public Sub Prepare(ByVal designShape As Shape, ByVal cutShape As Shape, _
                   ByVal settings As AASettings)
    Set mDesignShape = designShape
    Set mCutShape = cutShape
    Set mSettings = settings
    Set mVariants = CreateObject("Scripting.Dictionary")
    Set mMinimumSteps = CreateObject("Scripting.Dictionary")
    Set mMediumDesignGeometry = Nothing
    mTriangle = False
    Set mDesignGeometry = New AAGeometry
    Set mCutGeometry = Nothing
    If Not cutShape Is Nothing Then
        ' Minimum memakai Cut Line sebagai acuan; Design hanya membutuhkan ukuran.
        mDesignGeometry.CaptureBounds designShape
        Set mDesignBoundary = mDesignGeometry
        Set mCutGeometry = New AAGeometry
        mCutGeometry.Capture cutShape
        mSharpRectangle = mCutGeometry.IsSharpRectangle
        mLayoutSharpRectangle = mSharpRectangle
        mRectangle = mLayoutSharpRectangle Or (cutShape.Type = cdrRectangleShape)
        If Not mRectangle Then mRectangle = mCutGeometry.IsRoundedRectangle
        mTriangle = mCutGeometry.IsSharpTriangle
        If Not mTriangle Then mTriangle = mCutGeometry.IsRoundedTriangle
        If Not mSharpRectangle Then ValidateCutSize
    Else
        mDesignGeometry.Capture designShape
        Set mDesignBoundary = mDesignGeometry.OuterBoundary
        Set mDesignGeometry = mDesignBoundary
        mSharpRectangle = mDesignBoundary.IsSharpRectangle
        mLayoutSharpRectangle = mSharpRectangle
        mRectangle = mSharpRectangle Or (designShape.Type = cdrRectangleShape)
        If Not mRectangle Then mRectangle = mDesignBoundary.IsRoundedRectangle
        mTriangle = mDesignBoundary.IsSharpTriangle
        If Not mTriangle Then mTriangle = mDesignBoundary.IsRoundedTriangle
    End If
End Sub

Public Sub GetMinimumSteps(ByVal layoutMode As Long, ByRef horizontalStep As Double, _
                           ByRef verticalStep As Double)
    Dim cached As Variant, orientation As AARotationVariant, origin As AAPlacement
    Dim h As Double, v As Double, key As String, rowOffset As Double
    GetSteps layoutMode, horizontalStep, verticalStep
    If mRectangle Then Exit Sub
    key = CStr(layoutMode)
    If mMinimumSteps.Exists(key) Then
        cached = mMinimumSteps(key)
        horizontalStep = CDbl(cached(0)): verticalStep = CDbl(cached(1))
        Exit Sub
    End If
    GetBoundaryGaps layoutMode, h, v
    Set orientation = RotationVariant(0#)
    Set origin = orientation.MakePlacement(0, -1, 0, 0#, orientation.Height, _
        MaxNonnegative(h), MaxNonnegative(v))
    ' Persen tetap langkah Top/Left; nilai negatif mempertahankan kontrak lama.
    If h >= 0# And Not (mSettings.HasHorizontalGap And mSettings.HorizontalPercent) Then
        horizontalStep = ContactPitch(orientation, origin, True, horizontalStep, 0#, h, v)
    End If
    ' ZigZag kosong tetap 88%. Gap literal V memakai kontak dengan kedua tetangga.
    If v >= 0# And Not (mSettings.HasVerticalGap And mSettings.VerticalPercent) Then
        If layoutMode = AA_MODE_STRAIGHT Or mSettings.HasVerticalGap Then
            If layoutMode = AA_MODE_ZIGZAG Then rowOffset = horizontalStep / 2#
            verticalStep = ContactPitch(orientation, origin, False, _
                verticalStep, rowOffset, h, v)
        End If
    End If
    mMinimumSteps.Add key, Array(horizontalStep, verticalStep)
End Sub

Private Function MaxNonnegative(ByVal value As Double) As Double
    If value > 0# Then MaxNonnegative = value
End Function

Private Function ContactPitch(ByVal orientation As AARotationVariant, _
                              ByVal origin As AAPlacement, ByVal horizontal As Boolean, _
                              ByVal initial As Double, ByVal rowOffset As Double, _
                              ByVal h As Double, ByVal v As Double) As Double
    Dim safe As Double, trial As Double, blocked As Double, scanStep As Double
    Dim referenceSize As Double, mid As Double, lowLimit As Double
    referenceSize = orientation.ReferenceHeight
    If horizontal Then referenceSize = orientation.ReferenceWidth
    safe = referenceSize + MaxNonnegative(h)
    If v > h Then safe = referenceSize + MaxNonnegative(v)
    If initial > safe Then safe = initial
    safe = safe + 2# * orientation.Contour.MaxChordError + AA_SIZE_TOLERANCE_MM
    ' Pertahankan langkah lama yang sudah aman, termasuk area yang muat tepat.
    If ContactFits(orientation, origin, horizontal, initial, rowOffset, h, v) Then safe = initial
    scanStep = referenceSize / AA_CONTACT_SCAN_DIVISIONS
    If scanStep < 1# Then scanStep = 1#
    lowLimit = AA_CONTACT_TOLERANCE_MM
    If initial < lowLimit Then lowLimit = initial
    Do
        trial = safe - scanStep
        If trial < lowLimit Then trial = lowLimit
        If ContactFits(orientation, origin, horizontal, trial, rowOffset, h, v) Then
            safe = trial
            If trial = lowLimit Then Exit Do
        Else
            blocked = trial
            Do While safe - blocked > AA_CONTACT_TOLERANCE_MM
                mid = (safe + blocked) / 2#
                If ContactFits(orientation, origin, horizontal, mid, rowOffset, h, v) Then
                    safe = mid
                Else
                    blocked = mid
                End If
            Loop
            Exit Do
        End If
    Loop
    ContactPitch = safe
End Function

Private Function ContactFits(ByVal orientation As AARotationVariant, _
                             ByVal origin As AAPlacement, ByVal horizontal As Boolean, _
                             ByVal pitch As Double, ByVal rowOffset As Double, _
                             ByVal h As Double, ByVal v As Double) As Boolean
    Dim candidate As AAPlacement, x As Double, y As Double, side As Long
    Dim count As Long
    count = 1
    If Not horizontal And rowOffset <> 0# Then count = 3
    For side = 1 To count
        x = 0#: y = orientation.Height
        If horizontal Then
            x = pitch
        ElseIf side = 3 Then
            y = y - 2# * pitch
        Else
            x = rowOffset
            If side = 2 Then x = -rowOffset
            y = y - pitch
        End If
        If Not orientation.HasClearanceAt(x, y, origin, MaxNonnegative(h), _
            MaxNonnegative(v)) Then Exit Function
        Set candidate = orientation.MakePlacement(0, -1, 0, x, y, _
            MaxNonnegative(h), MaxNonnegative(v))
        If candidate.CollisionFootprint.Overlaps(origin.CollisionFootprint) Then Exit Function
    Next side
    ContactFits = True
End Function

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
    Dim h As Double, v As Double, orientation As AARotationVariant

    Set footprint = New AAGeometry
    footprint.AddPlaced mDesignGeometry, leftX, topY
    If Not mCutGeometry Is Nothing Then
        cutLeft = leftX + (DesignWidth - CutWidth) / 2#
        cutTop = topY - (DesignHeight - CutHeight) / 2#
        footprint.AddBounds cutLeft, cutTop - CutHeight, cutLeft + CutWidth, cutTop
    End If
    Set result = New AAPlacement
    result.Configure designIndex, cutIndex, areaIndex, leftX, topY, _
        cutLeft, cutTop, footprint
    If Not mCutGeometry Is Nothing Then _
        result.SetCollisionFootprint mCutGeometry.Placed(cutLeft, cutTop)
    GetBoundaryGaps AA_MODE_STRAIGHT, h, v
    Set orientation = RotationVariant(0#)
    result.SetContourData orientation.Contour, h, v, _
        Not (mSettings.HasHorizontalGap And mSettings.HorizontalPercent) And _
        Not (mSettings.HasVerticalGap And mSettings.VerticalPercent) And h >= 0# And v >= 0#
    Set MakePlacement = result
End Function

Public Function SameDesignGeometry(ByVal other As AAModel) As Boolean
    Dim first As AAGeometry, second As AAGeometry
    ' Jalur eksplisit ini tetap membandingkan artwork; planner Minimum tidak memakainya.
    Set first = New AAGeometry
    Set second = New AAGeometry
    first.Capture mDesignShape
    second.Capture other.DesignShape
    SameDesignGeometry = first.SameGeometry(second)
End Function

Public Function CanShareCutWith(ByVal other As AAModel) As Boolean
    If Not mCutGeometry Is Nothing Then
        If Not other.CutGeometry Is Nothing Then
            ' Ukuran Design sudah divalidasi; artwork bukan acuan Minimum.
            CanShareCutWith = mCutGeometry.SameGeometry(other.CutGeometry)
        End If
    ElseIf mSharpRectangle Or other.IsSharpRectangle Then
        CanShareCutWith = mDesignBoundary.SameGeometry(other.DesignBoundary)
    Else
        CanShareCutWith = mDesignBoundary.SameForm(other.DesignBoundary)
    End If
End Function

Friend Property Get CutGeometry() As AAGeometry
    Set CutGeometry = mCutGeometry
End Property

Friend Property Get DesignShape() As Shape
    Set DesignShape = mDesignShape
End Property

Friend Property Get DesignBoundary() As AAGeometry
    Set DesignBoundary = mDesignBoundary
End Property

Public Property Get IsRectangle() As Boolean
    IsRectangle = mRectangle
End Property

Public Property Get IsTriangle() As Boolean
    IsTriangle = mTriangle
End Property

Public Function RotationVariant(ByVal angle As Double) As AARotationVariant
    Dim result As AARotationVariant
    Dim design As AAGeometry, collision As AAGeometry
    If mVariants.Exists(CStr(angle)) Then
        Set RotationVariant = mVariants(CStr(angle))
        Exit Function
    End If
    If Not mCutGeometry Is Nothing And angle / 90# <> Int(angle / 90#) Then
        ' Pada sudut miring, rotasi bounding box saja terlalu besar untuk bentuk
        ' non-persegi. Baca artwork sekali untuk bounds rotasi, bukan collision.
        If mMediumDesignGeometry Is Nothing Then
            Set mMediumDesignGeometry = New AAGeometry
            mMediumDesignGeometry.Capture mDesignShape
        End If
        Set design = mMediumDesignGeometry.Placed(-DesignWidth / 2#, DesignHeight / 2#)
    Else
        Set design = mDesignGeometry.Placed(-DesignWidth / 2#, DesignHeight / 2#)
    End If
    Set design = design.Rotated(angle, 0#, 0#)
    If mCutGeometry Is Nothing Then
        Set collision = design
    Else
        Set collision = mCutGeometry.Placed(-CutWidth / 2#, CutHeight / 2#)
        Set collision = collision.Rotated(angle, 0#, 0#)
    End If
    Set result = New AARotationVariant
    result.Configure angle, design, collision
    mVariants.Add CStr(angle), result
    Set RotationVariant = result
End Function

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
