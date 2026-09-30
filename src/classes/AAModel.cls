Option Explicit

Private Const AA_DEFAULT_STRAIGHT_GAP_MM As Double = 1#
Private Const AA_DEFAULT_ZIGZAG_VERTICAL_FACTOR As Double = 0.88

Private mMarkerWidthMM As Double
Private mMarkerHeightMM As Double
Private mPositions As Collection
Private mLayoutMode As Long

Public Sub Prepare(ByVal layoutMode As Long, ByVal shapeWidth As Double, _
                   ByVal shapeHeight As Double, ByVal settings As AASettings)
    Dim straightPositions As Collection
    Dim zigZagPositions As Collection

    mMarkerWidthMM = settings.AreaWidthMM
    mMarkerHeightMM = settings.AreaHeightMM

    If layoutMode = AA_MODE_AUTO Then
        Set straightPositions = BuildPositions(AA_MODE_STRAIGHT, shapeWidth, _
            shapeHeight, settings)
        Set zigZagPositions = BuildPositions(AA_MODE_ZIGZAG, shapeWidth, _
            shapeHeight, settings)
        If zigZagPositions.Count > straightPositions.Count Then
            mLayoutMode = AA_MODE_ZIGZAG
            Set mPositions = zigZagPositions
        Else
            mLayoutMode = AA_MODE_STRAIGHT
            Set mPositions = straightPositions
        End If
    Else
        mLayoutMode = layoutMode
        Set mPositions = BuildPositions(layoutMode, shapeWidth, shapeHeight, settings)
    End If
End Sub

Public Property Get LayoutMode() As Long
    LayoutMode = mLayoutMode
End Property

Public Property Get MarkerWidthMM() As Double
    MarkerWidthMM = mMarkerWidthMM
End Property

Public Property Get MarkerHeightMM() As Double
    MarkerHeightMM = mMarkerHeightMM
End Property

Public Property Get Positions() As Collection
    Set Positions = mPositions
End Property

Public Property Get PositionCount() As Long
    If Not mPositions Is Nothing Then PositionCount = mPositions.Count
End Property

Private Function BuildPositions(ByVal layoutMode As Long, ByVal shapeWidth As Double, _
                                ByVal shapeHeight As Double, ByVal settings As AASettings) _
                                As Collection
    Dim horizontalStep As Double
    Dim verticalStep As Double

    Select Case layoutMode
        Case AA_MODE_DENSE
            horizontalStep = shapeWidth
            verticalStep = shapeHeight
        Case AA_MODE_STRAIGHT, AA_MODE_ZIGZAG
            horizontalStep = shapeWidth + AA_DEFAULT_STRAIGHT_GAP_MM
            If layoutMode = AA_MODE_STRAIGHT Then
                verticalStep = shapeHeight + AA_DEFAULT_STRAIGHT_GAP_MM
            Else
                verticalStep = shapeHeight * AA_DEFAULT_ZIGZAG_VERTICAL_FACTOR
            End If
            If settings.HasHorizontalGap Then
                If settings.HorizontalPercent Then
                    horizontalStep = shapeWidth * settings.HorizontalGap / 100#
                Else
                    horizontalStep = shapeWidth + settings.HorizontalGap
                End If
            End If
            If settings.HasVerticalGap Then
                If settings.VerticalPercent Then
                    verticalStep = shapeHeight * settings.VerticalGap / 100#
                Else
                    verticalStep = shapeHeight + settings.VerticalGap
                End If
            End If
        Case Else
            Err.Raise vbObjectError + 5203, "AAModel", "Mode penyusunan tidak dikenal."
    End Select

    Set BuildPositions = AAPlanPositions(shapeWidth, shapeHeight, mMarkerWidthMM, _
        mMarkerHeightMM, layoutMode, horizontalStep, verticalStep)
End Function
