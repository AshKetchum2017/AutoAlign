Option Explicit

Private Const AA_GROUP_GAP_MM As Double = 50#

Public Function Build(ByVal plan As AALayoutPlan, ByRef sourceShapes() As Shape, _
                      ByVal entryCount As Long, ByVal settings As AASettings, _
                      ByVal pageCenterX As Double, ByVal pageCenterY As Double) As ShapeRange
    Dim placements As Collection
    Dim placement As AAPlacement
    Dim designs() As Shape
    Dim cuts() As Shape
    Dim used() As Boolean
    Dim i As Long
    Dim j As Long
    Dim sourceBuckets As Object
    Dim parentBuckets As Object
    Dim sourceOrder As Collection
    Dim parentOrder As Collection
    Dim sourceBucket As AAGroupBucket
    Dim parentBucket As AAGroupBucket
    Dim key As Variant
    Dim result As ShapeRange
    Dim cutParent As AAGroupBucket
    Dim designParent As AAGroupBucket

    Set placements = plan.Placements
    ReDim designs(1 To placements.Count)
    ReDim cuts(1 To placements.Count)
    ReDim used(0 To entryCount - 1)

    For i = 1 To placements.Count
        Set placement = placements(i)
        Set designs(i) = TakeShape(sourceShapes(placement.DesignIndex), _
            used(placement.DesignIndex))
        PlaceShape designs(i), placement.DesignLeft, placement.DesignTop, _
            placement.AreaIndex, settings, pageCenterX, pageCenterY

        If placement.CutIndex >= 0 Then
            Set cuts(i) = TakeShape(sourceShapes(placement.CutIndex), _
                used(placement.CutIndex))
            PlaceShape cuts(i), placement.CutLeft, placement.CutTop, _
                placement.AreaIndex, settings, pageCenterX, pageCenterY
        End If
    Next i

    ' Pertahankan layer sumber; koreksi hanya pasangan dengan urutan terbalik.
    For i = 1 To placements.Count
        If Not cuts(i) Is Nothing Then AAOrderPlacementPair designs(i), cuts(i)
    Next i

    Set sourceBuckets = CreateObject("Scripting.Dictionary")
    Set sourceOrder = New Collection
    For i = 1 To placements.Count
        Set placement = placements(i)
        AddToBucket sourceBuckets, sourceOrder, placement.AreaIndex, _
            placement.DesignIndex, AA_ROLE_DESIGN, designs(i)
        If Not cuts(i) Is Nothing Then _
            AddToBucket sourceBuckets, sourceOrder, placement.AreaIndex, _
                placement.CutIndex, AA_ROLE_CUT_LINE, cuts(i)
    Next i

    Set parentBuckets = CreateObject("Scripting.Dictionary")
    Set parentOrder = New Collection
    For Each key In sourceOrder
        Set sourceBucket = sourceBuckets(CStr(key))
        sourceBucket.Finish True
        AddToBucket parentBuckets, parentOrder, sourceBucket.AreaIndex, _
            -1, sourceBucket.RoleLabel, sourceBucket.Result
    Next key

    Set result = New ShapeRange
    For Each key In parentOrder
        Set parentBucket = parentBuckets(CStr(key))
        parentBucket.Finish False
        result.Add parentBucket.Result
    Next key

    ' Semua Cut Line di depan semua Design pada layer yang sama.
    For Each key In parentOrder
        Set cutParent = parentBuckets(CStr(key))
        If cutParent.RoleLabel = AA_ROLE_CUT_LINE Then
            For j = 1 To parentOrder.Count
                Set designParent = parentBuckets(CStr(parentOrder(j)))
                If designParent.RoleLabel = AA_ROLE_DESIGN Then
                    If cutParent.Result.Layer.Index = designParent.Result.Layer.Index Then
                        If Not cutParent.Result.OrderIsInFrontOf(designParent.Result) Then _
                            cutParent.Result.OrderFrontOf designParent.Result
                    End If
                End If
            Next j
        End If
    Next key

    Set Build = result
End Function

Private Function TakeShape(ByVal sourceShape As Shape, ByRef alreadyUsed As Boolean) As Shape
    If alreadyUsed Then
        Set TakeShape = sourceShape.Duplicate(0#, 0#)
    Else
        Set TakeShape = sourceShape
        alreadyUsed = True
    End If
End Function

Private Sub PlaceShape(ByVal targetShape As Shape, ByVal localLeft As Double, _
                       ByVal localTop As Double, ByVal areaIndex As Long, _
                       ByVal settings As AASettings, ByVal pageCenterX As Double, _
                       ByVal pageCenterY As Double)
    targetShape.LeftX = pageCenterX - settings.AreaWidthMM / 2# + _
        areaIndex * (settings.AreaWidthMM + AA_GROUP_GAP_MM) + localLeft
    targetShape.TopY = pageCenterY - settings.AreaHeightMM / 2# + localTop
End Sub

Private Sub AddToBucket(ByVal buckets As Object, ByVal keyOrder As Collection, _
                        ByVal areaIndex As Long, ByVal entryIndex As Long, _
                        ByVal roleLabel As String, ByVal member As Shape)
    Dim key As String
    Dim bucket As AAGroupBucket

    key = CStr(areaIndex) & "|" & CStr(entryIndex) & "|" & roleLabel & _
        "|" & CStr(member.Layer.Index)
    If Not buckets.Exists(key) Then
        Set bucket = New AAGroupBucket
        bucket.Configure areaIndex, entryIndex, roleLabel, member.Layer
        buckets.Add key, bucket
        keyOrder.Add key
    Else
        Set bucket = buckets(key)
    End If
    bucket.Add member
End Sub
