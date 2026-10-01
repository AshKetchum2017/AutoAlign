Option Explicit

Private Const AA_DEFAULT_ZIGZAG_VERTICAL_FACTOR As Double = 0.88

Private mDesignShape As Shape
Private mCutShape As Shape
Private mDesignGeometry As AAGeometry
Private mCutGeometry As AAGeometry
Private mSettings As AASettings
Private mSharpRectangle As Boolean

Public Sub Prepare(ByVal designShape As Shape, ByVal cutShape As Shape, _
                   ByVal settings As AASettings)
    Set mDesignShape = designShape
    Set mCutShape = cutShape
    Set mSettings = settings
    Set mDesignGeometry = New AAGeometry
    mDesignGeometry.Capture designShape
    mSharpRectangle = mDesignGeometry.IsSharpRectangle
    If Not cutShape Is Nothing Then
        Set mCutGeometry = New AAGeometry
        mCutGeometry.Capture cutShape
    End If
End Sub

Public Sub GetSteps(ByVal layoutMode As Long, ByRef horizontalStep As Double, _
                    ByRef verticalStep As Double)
    Dim sharpPreset As Boolean
    If layoutMode <> AA_MODE_STRAIGHT And layoutMode <> AA_MODE_ZIGZAG Then Err.Raise _
        vbObjectError + 5203, "AAModel", "Mode penyusunan tidak dikenal."

    sharpPreset = mSharpRectangle And mSettings.ModelMode <> AA_MODEL_CUSTOM
    horizontalStep = DesignWidth + mSettings.DefaultHorizontalGapMM
    If sharpPreset Then horizontalStep = DesignWidth
    If layoutMode = AA_MODE_STRAIGHT Then
        verticalStep = DesignHeight + mSettings.DefaultVerticalGapMM
    Else
        verticalStep = DesignHeight * AA_DEFAULT_ZIGZAG_VERTICAL_FACTOR
    End If
    If sharpPreset Then verticalStep = DesignHeight
    If mSettings.HasHorizontalGap Then
        If mSettings.HorizontalPercent Then
            horizontalStep = DesignWidth * mSettings.HorizontalGap / 100#
        Else
            horizontalStep = DesignWidth + mSettings.HorizontalGap
        End If
    End If
    If mSettings.HasVerticalGap Then
        If mSettings.VerticalPercent Then
            verticalStep = DesignHeight * mSettings.VerticalGap / 100#
        Else
            verticalStep = DesignHeight + mSettings.VerticalGap
        End If
    End If
    If horizontalStep <= 0# Or verticalStep <= 0# Then Err.Raise _
        vbObjectError + 5205, "AAModel", "Langkah penyusunan harus lebih besar dari nol."
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
    Set MakePlacement = result
End Function

Public Function SameDesignGeometry(ByVal other As AAModel) As Boolean
    SameDesignGeometry = mDesignGeometry.SameGeometry(other.DesignGeometry)
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

Public Property Get CutWidth() As Double
    If Not mCutGeometry Is Nothing Then CutWidth = mCutGeometry.Width
End Property

Public Property Get CutHeight() As Double
    If Not mCutGeometry Is Nothing Then CutHeight = mCutGeometry.Height
End Property
