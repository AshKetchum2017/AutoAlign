Option Explicit

Private Const AA_DEFAULT_MARKER_WIDTH_MM As Double = 320#
Private Const AA_DEFAULT_MARKER_HEIGHT_MM As Double = 470#
Private Const AA_DEFAULT_STRAIGHT_GAP_MM As Double = 1#
Private Const AA_DEFAULT_ZIGZAG_VERTICAL_FACTOR As Double = 0.88

Private mMarkerWidthMM As Double
Private mMarkerHeightMM As Double
Private mPositions As Collection

Public Sub Prepare(ByVal layoutMode As Long, ByVal shapeWidth As Double, _
                   ByVal shapeHeight As Double)
    mMarkerWidthMM = AA_DEFAULT_MARKER_WIDTH_MM
    mMarkerHeightMM = AA_DEFAULT_MARKER_HEIGHT_MM
    Set mPositions = AAPlanPositions(shapeWidth, shapeHeight, mMarkerWidthMM, _
        mMarkerHeightMM, layoutMode, AA_DEFAULT_STRAIGHT_GAP_MM, _
        AA_DEFAULT_ZIGZAG_VERTICAL_FACTOR)
End Sub

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
