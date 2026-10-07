Option Explicit

Private mAngle As Double
Private mDesign As AAGeometry
Private mCollision As AAGeometry
Private mBounds As AAGeometry
Private mContour As AAContour

Public Sub Configure(ByVal angle As Double, ByVal design As AAGeometry, _
                     ByVal collision As AAGeometry)
    mAngle = angle
    Set mDesign = design
    Set mCollision = collision
    Set mBounds = New AAGeometry
    mBounds.AddBounds design.LeftX, design.BottomY, design.RightX, design.TopY
    mBounds.AddBounds collision.LeftX, collision.BottomY, collision.RightX, collision.TopY
    Set mContour = New AAContour
    mContour.Capture collision
End Sub

Public Function MakePlacement(ByVal designIndex As Long, ByVal cutIndex As Long, _
                              ByVal areaIndex As Long, ByVal leftX As Double, _
                              ByVal topY As Double, ByVal gapH As Double, _
                              ByVal gapV As Double) As AAPlacement
    Dim result As AAPlacement
    Dim footprint As AAGeometry, collision As AAGeometry
    Dim dx As Double, dy As Double
    Dim perfStarted As Double
    If AADebugProfiling Then perfStarted = AADebugPerfBegin()
    dx = leftX - mBounds.LeftX: dy = topY - mBounds.TopY
    Set footprint = mBounds.Placed(leftX, topY)
    Set collision = mCollision.Placed(mCollision.LeftX + dx, mCollision.TopY + dy)
    Set result = New AAPlacement
    result.Configure designIndex, cutIndex, areaIndex, mDesign.LeftX + dx, _
        mDesign.TopY + dy, mCollision.LeftX + dx, mCollision.TopY + dy, footprint
    result.SetCollisionFootprint collision
    result.SetMediumData mAngle, dx, dy, mContour, gapH, gapV
    Set MakePlacement = result
    If AADebugProfiling Then AADebugPerfEnd AA_PERF_PLACEMENT, perfStarted
End Function

Public Property Get Width() As Double
    Width = mBounds.Width
End Property

Public Property Get Height() As Double
    Height = mBounds.Height
End Property

Public Property Get ReferenceWidth() As Double
    ReferenceWidth = mCollision.Width
End Property

Public Property Get ReferenceHeight() As Double
    ReferenceHeight = mCollision.Height
End Property

Public Property Get InsetX() As Double
    InsetX = mCollision.LeftX - mBounds.LeftX
End Property

Public Property Get InsetY() As Double
    InsetY = mBounds.TopY - mCollision.TopY
End Property

Public Property Get Angle() As Double
    Angle = mAngle
End Property

Public Property Get Contour() As AAContour
    Set Contour = mContour
End Property

' Uji gap numerik memakai kontur rotasi yang sudah disimpan. Kurva kandidat
' baru disalin setelah kandidat lolos; overlap tetap diuji dengan kurva asli.
Public Function HasClearanceAt(ByVal leftX As Double, ByVal topY As Double, _
                               ByVal other As AAPlacement, ByVal gapH As Double, _
                               ByVal gapV As Double) As Boolean
    Dim h As Double, v As Double
    h = gapH: v = gapV
    If other.GapH > h Then h = other.GapH
    If other.GapV > v Then v = other.GapV
    HasClearanceAt = mContour.ClearOf(other.Contour, _
        leftX - mBounds.LeftX, topY - mBounds.TopY, _
        other.CollisionFootprint.LeftX - other.Contour.LeftX, _
        other.CollisionFootprint.TopY - other.Contour.TopY, h, v)
End Function
