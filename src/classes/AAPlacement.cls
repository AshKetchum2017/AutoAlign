Option Explicit

Private mDesignIndex As Long
Private mCutIndex As Long
Private mAreaIndex As Long
Private mDesignLeft As Double
Private mDesignTop As Double
Private mCutLeft As Double
Private mCutTop As Double
Private mFootprint As AAGeometry

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
End Sub

Public Sub Translate(ByVal deltaX As Double, ByVal deltaY As Double)
    Set mFootprint = mFootprint.Placed(mFootprint.LeftX + deltaX, _
        mFootprint.TopY + deltaY)
    mDesignLeft = mDesignLeft + deltaX
    mDesignTop = mDesignTop + deltaY
    mCutLeft = mCutLeft + deltaX
    mCutTop = mCutTop + deltaY
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
