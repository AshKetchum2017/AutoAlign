Attribute VB_Name = "AADebug"
Option Explicit

' Immediate Window: AADebugOn (ringkas), AADebugOn True (detail), AADebugOff.
' Tidak disimpan: reset project mengembalikan debug ke kondisi nonaktif.
Private mEnabled As Boolean
Private mDetailed As Boolean
Private mSlotsActive As Boolean
Private mCandidate As String
Private mBoundsRejected As Long, mGapRejected As Long
Private mMarkRejected As Long, mOverlapRejected As Long

Public Sub AADebugOn(Optional ByVal detailed As Boolean = False)
    mEnabled = True
    mDetailed = detailed
    mSlotsActive = False
    Debug.Print "[AA][Debug] ON; detailed=" & CStr(detailed)
End Sub

Public Sub AADebugOff()
    mEnabled = False
    mDetailed = False
    mSlotsActive = False
    Debug.Print "[AA][Debug] OFF"
End Sub

Public Function AADebugEnabled() As Boolean
    AADebugEnabled = mEnabled
End Function

Public Function AADebugDetailsEnabled() As Boolean
    AADebugDetailsEnabled = mEnabled And mDetailed
End Function

Public Sub AADebugWrite(ByVal section As String, ByVal message As String)
    If Not mEnabled Then Exit Sub
    On Error Resume Next
    Debug.Print "[AA][" & section & "] " & message
    Err.Clear
End Sub

Public Sub AADebugProcessStart(ByVal layoutMode As Long, ByVal modelMode As Long)
    Dim layoutName As String, modelName As String
    If Not mEnabled Then Exit Sub
    mSlotsActive = False
    layoutName = "Minimum"
    If layoutMode = AA_MODE_MEDIUM Then layoutName = "Medium"
    modelName = "Custom"
    If modelMode = AA_MODEL_KISS_A Then modelName = "KissA"
    If modelMode = AA_MODEL_DIE_A Then modelName = "DieA"
    AADebugWrite "Process", layoutName & " / " & modelName
End Sub

Public Sub AADebugCandidateStart(ByVal label As String)
    mSlotsActive = mEnabled And mDetailed
    If Not mSlotsActive Then Exit Sub
    mCandidate = label
    mBoundsRejected = 0: mGapRejected = 0
    mMarkRejected = 0: mOverlapRejected = 0
End Sub

Public Sub AADebugCandidateEnd()
    If Not mSlotsActive Then Exit Sub
    mSlotsActive = False
    AADebugWrite "Slots", mCandidate & "; rejected bounds=" & CStr(mBoundsRejected) & _
        "; gap=" & CStr(mGapRejected) & "; mark=" & CStr(mMarkRejected) & _
        "; overlap=" & CStr(mOverlapRejected)
End Sub

Public Sub AADebugRejectSlot(ByVal reason As String)
    If Not mSlotsActive Then Exit Sub
    Select Case reason
        Case "bounds": mBoundsRejected = mBoundsRejected + 1
        Case "gap": mGapRejected = mGapRejected + 1
        Case "mark": mMarkRejected = mMarkRejected + 1
        Case "overlap": mOverlapRejected = mOverlapRejected + 1
    End Select
End Sub

' Saat tracing detail aktif, hanya tiga contoh per kandidat yang membaca
' ulang kurva. Counter tetap mencakup semua penolakan dalam kandidat itu.
Public Sub AADebugGapRejected(ByVal orientation As AARotationVariant, _
                              ByVal x As Double, ByVal y As Double, _
                              ByVal blocker As AAPlacement, _
                              ByVal h As Double, ByVal v As Double)
    Dim candidate As AAPlacement
    If Not mSlotsActive Then Exit Sub
    mGapRejected = mGapRejected + 1
    If mGapRejected > 3 Then Exit Sub
    On Error GoTo TraceFailed
    Set candidate = orientation.MakePlacement(0, -1, 0, x, y, h, v)
    AADebugWrite "Slot", "gap angle=" & CStr(orientation.Angle) & _
        "; x=" & Format$(x, "0.000000000") & "; y=" & Format$(y, "0.000000000") & _
        "; blocker angle=" & CStr(blocker.RotationDelta) & _
        "; placed clearance=" & CStr(candidate.HasClearance(blocker)) & _
        "; curve shift error=" & Format$(candidate.CollisionFootprint.LeftX - _
            x - orientation.InsetX, "0.000000000") & "/" & _
        Format$(candidate.CollisionFootprint.TopY - y + orientation.InsetY, "0.000000000")
TraceFailed:
    Err.Clear
End Sub

Public Sub AADebugOverlapRejected(ByVal candidate As AAPlacement, _
                                  ByVal blocker As AAPlacement)
    Dim first As Curve, second As Curve, joined As Curve
    If Not mSlotsActive Then Exit Sub
    mOverlapRejected = mOverlapRejected + 1
    If mOverlapRejected > 3 Then Exit Sub
    On Error GoTo TraceFailed
    AADebugWrite "Slot", "overlap angle=" & CStr(candidate.RotationDelta) & _
        "; x=" & Format$(candidate.Footprint.LeftX, "0.000000000") & _
        "; y=" & Format$(candidate.Footprint.TopY, "0.000000000") & _
        "; blocker angle=" & CStr(blocker.RotationDelta)
    If candidate.CollisionFootprint.CurveCount <> 1 Or _
       blocker.CollisionFootprint.CurveCount <> 1 Then Exit Sub
    Set first = candidate.CollisionFootprint.CurveAt(1)
    Set second = blocker.CollisionFootprint.CurveAt(1)
    Set joined = first.WeldWith(second)
    AADebugWrite "Slot", "union area difference=" & _
        Format$(Abs(first.Area) + Abs(second.Area) - Abs(joined.Area), "0.000000000000")
TraceFailed:
    Err.Clear
End Sub

Public Sub AADebugPlan(ByVal label As String, ByVal plan As AALayoutPlan)
    Dim widthMM As Double, heightMM As Double, boundsArea As Double
    Dim collisionWidth As Double, collisionHeight As Double, collisionArea As Double
    If Not mEnabled Then Exit Sub
    On Error GoTo TraceFailed
    If plan Is Nothing Then
        AADebugWrite "Plan", label & ": rejected (no valid plan)"
        Exit Sub
    End If
    PlanBounds plan, False, widthMM, heightMM, boundsArea
    PlanBounds plan, True, collisionWidth, collisionHeight, collisionArea
    AADebugWrite "Plan", label & ": areas=" & CStr(plan.AreaCount) & _
        "; objects=" & CStr(plan.DesignCount) & _
        "; regular=" & CStr(plan.HasRegularPattern) & _
        "; max footprint=" & Format$(widthMM, "0.000") & "x" & Format$(heightMM, "0.000") & _
        " mm; collision area=" & Format$(collisionArea, "0.000") & " mm2"
TraceFailed:
    Err.Clear
End Sub

Public Sub AADebugSettings(ByVal settings As AASettings)
    Dim h As String, v As String
    If Not mEnabled Then Exit Sub
    On Error GoTo TraceFailed
    h = "default": v = "default"
    If settings.HasHorizontalGap Then h = CStr(settings.HorizontalGap)
    If settings.HasVerticalGap Then v = CStr(settings.VerticalGap)
    If settings.HorizontalPercent Then h = h & "%"
    If settings.VerticalPercent Then v = v & "%"
    AADebugWrite "Input", "area=" & CStr(settings.AreaWidthMM) & "x" & _
        CStr(settings.AreaHeightMM) & " mm; gap H=" & h & "; V=" & v
TraceFailed:
    Err.Clear
End Sub

Public Sub AADebugModel(ByVal index As Long, ByVal model As AAModel, ByVal quantity As Long)
    Dim orientation As AARotationVariant, h As Double, v As Double
    If Not AADebugDetailsEnabled Then Exit Sub
    On Error GoTo TraceFailed
    model.GetBoundaryGaps AA_MODE_STRAIGHT, h, v
    Set orientation = model.RotationVariant(0#)
    AADebugWrite "Model", "index=" & CStr(index) & "; quantity=" & CStr(quantity) & _
        "; size=" & Format$(orientation.Width, "0.000000") & "x" & _
        Format$(orientation.Height, "0.000000") & "; edges=" & CStr(orientation.Contour.Edges.Count) & _
        "; rectangle=" & CStr(model.IsRectangle) & "; triangle=" & CStr(model.IsTriangle) & _
        "; straight gap=" & Format$(h, "0.000000") & "/" & Format$(v, "0.000000")
TraceFailed:
    Err.Clear
End Sub

Private Sub PlanBounds(ByVal plan As AALayoutPlan, ByVal collision As Boolean, _
                       ByRef widthMM As Double, ByRef heightMM As Double, _
                       ByRef boundsArea As Double)
    Dim lefts() As Double, bottoms() As Double, rights() As Double, tops() As Double
    Dim seen() As Boolean, i As Long
    Dim placed As AAPlacement, box As AAGeometry
    If plan.AreaCount = 0 Then Exit Sub
    ReDim lefts(0 To plan.AreaCount - 1): ReDim bottoms(0 To plan.AreaCount - 1)
    ReDim rights(0 To plan.AreaCount - 1): ReDim tops(0 To plan.AreaCount - 1)
    ReDim seen(0 To plan.AreaCount - 1)
    For Each placed In plan.Placements
        i = placed.AreaIndex
        If collision Then Set box = placed.CollisionFootprint Else Set box = placed.Footprint
        If Not seen(i) Then
            lefts(i) = box.LeftX: bottoms(i) = box.BottomY
            rights(i) = box.RightX: tops(i) = box.TopY
            seen(i) = True
        Else
            If box.LeftX < lefts(i) Then lefts(i) = box.LeftX
            If box.BottomY < bottoms(i) Then bottoms(i) = box.BottomY
            If box.RightX > rights(i) Then rights(i) = box.RightX
            If box.TopY > tops(i) Then tops(i) = box.TopY
        End If
    Next placed
    For i = 0 To plan.AreaCount - 1
        If seen(i) Then
            If rights(i) - lefts(i) > widthMM Then widthMM = rights(i) - lefts(i)
            If tops(i) - bottoms(i) > heightMM Then heightMM = tops(i) - bottoms(i)
            boundsArea = boundsArea + (rights(i) - lefts(i)) * (tops(i) - bottoms(i))
        End If
    Next i
End Sub
